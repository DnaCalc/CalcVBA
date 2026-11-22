//! Built-in `VBA.Collection` instance state.
//!
//! The data model lives here in `calcvba-runtime` so every consumer — vm3 and the
//! future JIT — shares ONE implementation. Each `New Collection` instance's ordered,
//! optionally string-keyed contents ride the refcounted object box itself (the
//! `native_state` slot on `CompatObjectBase`, reached via
//! [`crate::object_ref::ObjectRef::with_native_collection`]), not a VM-side map. Because a
//! `Set c2 = c1` clones the `ObjectRef` (same box), both names see the same
//! `CollectionData` — VBA reference semantics fall out for free — and box teardown
//! reclaims it, so there is no leak.
//!
//! This is the data model only; the shared dispatch site
//! ([`crate::collection`] has none — see `calcvba_eval::collection::dispatch_collection`)
//! maps [`CollectionError`] onto VBA run-time error numbers and builds
//! [`Selector`]s from the call arguments. Keys are case-insensitive (folded to Unicode
//! lowercase by [`fold_key`]).

use crate::{ObjectRef, ObjectRefExt, Variant};
use std::cell::Cell;
use std::rc::{Rc, Weak};

/// Selects an existing element (for `Item`/`Remove`) or an insertion anchor (for
/// `Add before`/`after`): a 1-based index, or a case-insensitive string key
/// (stored folded).
pub enum Selector {
    Index(i32),
    Key(String),
}

/// A collection operation failure, mapped to a VBA run-time error number at the
/// dispatch site.
#[derive(Debug, PartialEq, Eq)]
pub enum CollectionError {
    /// Index out of range → error 9.
    NotFound,
    /// `Add` with a key already present → error 457.
    DuplicateKey,
    /// Empty collection, both insertion anchors, or an absent string key → error 5.
    BadArgument,
    /// `Add` given a key that is not a String → error 13.
    TypeMismatch,
    /// A required `Item`/`Remove` selector was omitted → error 449.
    ArgNotOptional,
    /// `Add` was called without an item → error 450.
    WrongArgumentCount,
}

impl CollectionError {
    pub fn vba_number(self) -> i32 {
        match self {
            Self::NotFound => 9,
            Self::DuplicateKey => 457,
            Self::BadArgument => 5,
            Self::TypeMismatch => 13,
            Self::ArgNotOptional => 449,
            Self::WrongArgumentCount => 450,
        }
    }
}

/// The case-insensitive form of a collection key: Unicode lower case, so `Ä` and `ä` are one key.
pub fn fold_key(key: &str) -> String {
    key.chars().flat_map(char::to_lowercase).collect()
}

/// A VBA `Collection`'s ordered contents. Indices are 1-based at the VBA surface.
#[derive(Debug, Default)]
pub struct CollectionData {
    entries: Vec<Entry>,
    cursors: Vec<Weak<Cell<usize>>>,
}

impl Clone for CollectionData {
    fn clone(&self) -> Self {
        Self {
            entries: self.entries.clone(),
            cursors: Vec::new(),
        }
    }
}

/// An owning `For Each` cursor over a live Collection.
pub struct CollectionCursor {
    collection: ObjectRef,
    next: Rc<Cell<usize>>,
}

impl CollectionCursor {
    pub fn new(collection: &ObjectRef) -> Option<Self> {
        let next = Rc::new(Cell::new(0));
        collection.with_native_collection(|data| {
            data.cursors.retain(|cursor| cursor.strong_count() != 0);
            data.cursors.push(Rc::downgrade(&next));
        })?;
        Some(Self {
            collection: collection.clone(),
            next,
        })
    }

    pub fn next_item(&self) -> Option<Variant> {
        self.collection.with_native_collection(|data| {
            let index = self.next.get();
            let value = data.entries.get(index)?.value.clone();
            self.next.set(index + 1);
            Some(value)
        })?
    }
}

#[derive(Debug, Clone)]
struct Entry {
    /// Folded (ASCII-lowercased) key, if the item was added with one.
    key: Option<String>,
    value: Variant,
}

impl CollectionData {
    /// `Collection.Count`.
    pub fn count(&self) -> i32 {
        self.entries.len() as i32
    }

    /// A snapshot of the values in collection order.
    pub fn values(&self) -> Vec<Variant> {
        self.entries.iter().map(|e| e.value.clone()).collect()
    }

    /// `Collection.Add item, [key], [before], [after]`. `key` is pre-folded. At
    /// most one of `before`/`after` may be given.
    pub fn add(
        &mut self,
        value: Variant,
        key: Option<String>,
        before: Option<Selector>,
        after: Option<Selector>,
    ) -> Result<(), CollectionError> {
        if before.is_some() && after.is_some() {
            return Err(CollectionError::BadArgument);
        }
        if let Some(k) = &key
            && self.position_of_key(k).is_some()
        {
            return Err(CollectionError::DuplicateKey);
        }
        let at = match (before, after) {
            (Some(sel), _) => self.resolve_index(&sel)?,
            (_, Some(sel)) => self.resolve_index(&sel)? + 1,
            _ => self.entries.len(),
        };
        if at < self.entries.len() {
            // An insertion before the pending item does not replace it.
            self.adjust_cursors(at, |next| next + 1);
        }
        self.entries.insert(at, Entry { key, value });
        Ok(())
    }

    /// `Collection.Item(indexOrKey)`.
    pub fn item(&self, sel: &Selector) -> Result<Variant, CollectionError> {
        let i = self.resolve_index(sel)?;
        Ok(self.entries[i].value.clone())
    }

    /// `Collection.Remove indexOrKey`.
    pub fn remove(&mut self, sel: &Selector) -> Result<(), CollectionError> {
        let i = self.resolve_index(sel)?;
        let removed = self.entries.remove(i);
        // The item after the removed one now occupies its position.
        self.adjust_cursors(i + 1, |next| next - 1);
        drop(removed);
        Ok(())
    }

    fn adjust_cursors(&mut self, first: usize, adjust: impl Fn(usize) -> usize) {
        self.cursors.retain(|cursor| {
            let Some(cursor) = cursor.upgrade() else {
                return false;
            };
            let next = cursor.get();
            if next >= first {
                cursor.set(adjust(next));
            }
            true
        });
    }

    /// Zero-based position of `sel`, or [`CollectionError::NotFound`].
    fn resolve_index(&self, sel: &Selector) -> Result<usize, CollectionError> {
        // Excel reports error5 for addressing an empty collection, including an Add anchor.
        if self.entries.is_empty() {
            return Err(CollectionError::BadArgument);
        }
        match sel {
            Selector::Index(i) => {
                let zero = i
                    .checked_sub(1)
                    .and_then(|z| usize::try_from(z).ok())
                    .ok_or(CollectionError::NotFound)?;
                (zero < self.entries.len())
                    .then_some(zero)
                    .ok_or(CollectionError::NotFound)
            }
            Selector::Key(k) => self.position_of_key(k).ok_or(CollectionError::BadArgument),
        }
    }

    fn position_of_key(&self, folded: &str) -> Option<usize> {
        self.entries
            .iter()
            .position(|e| e.key.as_deref() == Some(folded))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn i32_of(v: Result<Variant, CollectionError>) -> Option<i32> {
        v.ok().and_then(|v| v.as_i32())
    }

    #[test]
    fn positional_add_count_item_remove() {
        let mut c = CollectionData::default();
        assert_eq!(c.count(), 0);
        c.add(Variant::from_i32(10), None, None, None).unwrap();
        c.add(Variant::from_i32(20), None, None, None).unwrap();
        c.add(Variant::from_i32(30), None, None, None).unwrap();
        assert_eq!(c.count(), 3);
        assert_eq!(i32_of(c.item(&Selector::Index(1))), Some(10));
        assert_eq!(i32_of(c.item(&Selector::Index(3))), Some(30));
        assert_eq!(
            c.item(&Selector::Index(0)).unwrap_err(),
            CollectionError::NotFound
        );
        assert_eq!(
            c.item(&Selector::Index(4)).unwrap_err(),
            CollectionError::NotFound
        );
        c.remove(&Selector::Index(2)).unwrap(); // drop 20
        assert_eq!(c.count(), 2);
        assert_eq!(i32_of(c.item(&Selector::Index(2))), Some(30));
    }

    #[test]
    fn string_keys_are_case_insensitive_and_unique() {
        let mut c = CollectionData::default();
        c.add(Variant::from_i32(1), Some("alpha".into()), None, None)
            .unwrap();
        c.add(Variant::from_i32(2), Some("beta".into()), None, None)
            .unwrap();
        // Keyed lookup (the dispatch site folds the incoming key).
        assert_eq!(i32_of(c.item(&Selector::Key("alpha".into()))), Some(1));
        assert_eq!(i32_of(c.item(&Selector::Key("beta".into()))), Some(2));
        assert_eq!(
            c.item(&Selector::Key("missing".into())).unwrap_err(),
            CollectionError::BadArgument
        );
        // Duplicate key → 457.
        assert_eq!(
            c.add(Variant::from_i32(9), Some("alpha".into()), None, None)
                .unwrap_err(),
            CollectionError::DuplicateKey
        );
        // Remove by key.
        c.remove(&Selector::Key("alpha".into())).unwrap();
        assert_eq!(c.count(), 1);
        assert_eq!(i32_of(c.item(&Selector::Index(1))), Some(2));
    }

    #[test]
    fn before_and_after_position_the_insertion() {
        let mut c = CollectionData::default();
        c.add(Variant::from_i32(10), None, None, None).unwrap();
        c.add(Variant::from_i32(30), None, None, None).unwrap();
        // Insert 20 before index 2 → [10, 20, 30].
        c.add(Variant::from_i32(20), None, Some(Selector::Index(2)), None)
            .unwrap();
        // Insert 5 before index 1 → [5, 10, 20, 30].
        c.add(Variant::from_i32(5), None, Some(Selector::Index(1)), None)
            .unwrap();
        // Insert 40 after the last → [5, 10, 20, 30, 40].
        c.add(Variant::from_i32(40), None, None, Some(Selector::Index(4)))
            .unwrap();
        let vals: Vec<i32> = c.values().iter().filter_map(|v| v.as_i32()).collect();
        assert_eq!(vals, vec![5, 10, 20, 30, 40]);
    }

    #[test]
    fn after_a_key_inserts_following_it() {
        let mut c = CollectionData::default();
        c.add(Variant::from_i32(1), Some("a".into()), None, None)
            .unwrap();
        c.add(Variant::from_i32(3), Some("c".into()), None, None)
            .unwrap();
        c.add(
            Variant::from_i32(2),
            Some("b".into()),
            None,
            Some(Selector::Key("a".into())),
        )
        .unwrap();
        let vals: Vec<i32> = c.values().iter().filter_map(|v| v.as_i32()).collect();
        assert_eq!(vals, vec![1, 2, 3]);
    }

    #[test]
    fn both_before_and_after_is_a_bad_argument() {
        let mut c = CollectionData::default();
        c.add(Variant::from_i32(1), None, None, None).unwrap();
        assert_eq!(
            c.add(
                Variant::from_i32(2),
                None,
                Some(Selector::Index(1)),
                Some(Selector::Index(1)),
            )
            .unwrap_err(),
            CollectionError::BadArgument
        );
    }
}
