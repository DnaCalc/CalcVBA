use std::{cell::RefCell, collections::HashMap, rc::Rc};

use calcvba_com::DynamicCallArg;
use calcvba_runtime::{Variant, variant_to_vba_string};
use roxmltree::{Document, Node, NodeId};

use super::{BuiltinObject, StandardHostServices, argument, fault};
use crate::{HalResult, model::CapabilityId};

struct XmlSource {
    text: String,
    identities: RefCell<HashMap<NodeId, i32>>,
}

pub(in crate::adapters::standard) struct XmlElement {
    source: Rc<XmlSource>,
    id: NodeId,
    pub name: String,
    pub text: String,
}

impl XmlElement {
    pub fn parse(xml: &str) -> Option<Self> {
        let document = Document::parse(xml).ok()?;
        let source = Rc::new(XmlSource {
            text: xml.to_owned(),
            identities: RefCell::default(),
        });
        Some(Self::from_node(&source, document.root_element()))
    }

    fn from_node(source: &Rc<XmlSource>, node: Node<'_, '_>) -> Self {
        let name = source.text[node.range().start + 1..]
            .split([' ', '\t', '\r', '\n', '/', '>'])
            .next()
            .unwrap_or_default()
            .to_owned();
        let text: String = node
            .descendants()
            .filter(Node::is_text)
            .filter_map(|node| node.text())
            .collect();
        Self {
            source: source.clone(),
            id: node.id(),
            name,
            text: text.trim_matches([' ', '\t', '\r', '\n']).to_owned(),
        }
    }

    pub fn identity(&self, host: &StandardHostServices) -> HalResult<i32> {
        let mut identities = self.source.identities.borrow_mut();
        if let Some(&identity) = identities.get(&self.id) {
            return Ok(identity);
        }
        let identity = host
            .borrow_projection(CapabilityId::ComActivationDispatch, "create_object")?
            .allocate_handle();
        identities.insert(self.id, identity);
        Ok(identity)
    }

    pub fn select(
        &self,
        host: &StandardHostServices,
        args: &[DynamicCallArg],
        document_scope: bool,
    ) -> HalResult<Variant> {
        let query = variant_to_vba_string(argument(host, args, "queryString")?)
            .map_err(|_| fault(host, 13))?;
        let query = query.as_str();
        let parsed = Document::parse(&self.source.text).map_err(|_| fault(host, 438))?;
        let root = parsed.root();
        let context = if document_scope {
            root
        } else {
            parsed.get_node(self.id).ok_or_else(|| fault(host, 438))?
        };
        let (context, selector, descendants) = if let Some(path) = query.strip_prefix(".//") {
            (context, path, true)
        } else if let Some(path) = query.strip_prefix("//") {
            (root, path, true)
        } else if let Some(path) = query.strip_prefix('/') {
            (root, path, false)
        } else {
            (context, query.as_str(), false)
        };
        let mut names = selector.split('/');
        let first = names.next().unwrap_or_default();
        if !is_element_name(first) || !names.clone().all(is_element_name) {
            return Err(fault(host, 438));
        }

        let mut nodes: Vec<_> = if descendants {
            context
                .descendants()
                .skip(1)
                .filter(|&node| matches_name(node, first))
                .collect()
        } else {
            context
                .children()
                .filter(|&node| matches_name(node, first))
                .collect()
        };
        for name in names {
            nodes = nodes
                .into_iter()
                .flat_map(|node| node.children())
                .filter(|&node| matches_name(node, name))
                .collect();
        }
        nodes.first().map_or_else(
            || Ok(Variant::nothing()),
            |&node| BuiltinObject::Element(Self::from_node(&self.source, node)).retain(host),
        )
    }
}

fn is_element_name(name: &str) -> bool {
    name == "*" || {
        let mut chars = name.chars();
        chars.next().is_some_and(|c| c == '_' || c.is_alphabetic())
            && chars.all(|c| matches!(c, '_' | '-' | '.') || c.is_alphanumeric())
    }
}

fn matches_name(node: Node<'_, '_>, name: &str) -> bool {
    node.is_element()
        && (name == "*" || node.tag_name().namespace().is_none() && node.tag_name().name() == name)
}
