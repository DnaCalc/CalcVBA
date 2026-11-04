//! MS-OVBA 2.4.1 compressed containers. Copy tokens never cross chunk boundaries.
use crate::{Error, Result, invalid};

pub fn decompress(input: &[u8], limit: usize) -> Result<Vec<u8>> {
    if input.first() != Some(&1) {
        return Err(invalid("compressed-container signature"));
    }
    let mut at = 1usize;
    let mut output = Vec::new();
    while at < input.len() {
        let bytes = input
            .get(at..at + 2)
            .ok_or_else(|| invalid("chunk header"))?;
        let header = u16::from_le_bytes([bytes[0], bytes[1]]);
        if header & 0x7000 != 0x3000 {
            return Err(invalid("chunk signature"));
        }
        let end = at
            .checked_add(usize::from(header & 0xfff) + 3)
            .filter(|end| *end <= input.len())
            .ok_or_else(|| invalid("chunk size"))?;
        at += 2;
        let start = output.len();
        if header & 0x8000 == 0 {
            if end - at != 4096 {
                return Err(invalid("raw chunk size"));
            }
            if output.len().checked_add(4096).is_none_or(|n| n > limit) {
                return Err(Error::Limit("decompressed bytes"));
            }
            output.extend_from_slice(&input[at..end]);
            at = end;
            continue;
        }
        while at < end {
            let flags = input[at];
            at += 1;
            for bit in 0..8 {
                if at == end {
                    break;
                }
                if flags & (1 << bit) == 0 {
                    if output.len() == limit {
                        return Err(Error::Limit("decompressed bytes"));
                    }
                    if output.len() - start == 4096 {
                        return Err(invalid("expanded chunk size"));
                    }
                    output.push(input[at]);
                    at += 1;
                } else {
                    if end - at < 2 {
                        return Err(invalid("copy token"));
                    }
                    let token = u16::from_le_bytes([input[at], input[at + 1]]);
                    at += 2;
                    let size = output.len() - start;
                    let bits = offset_bits(size);
                    let mask = 0xffffu16 >> bits;
                    let length = usize::from(token & mask) + 3;
                    let distance = usize::from(token >> (16 - bits)) + 1;
                    if distance > size || size + length > 4096 {
                        return Err(invalid("copy distance or expanded chunk size"));
                    }
                    if output.len().checked_add(length).is_none_or(|n| n > limit) {
                        return Err(Error::Limit("decompressed bytes"));
                    }
                    for _ in 0..length {
                        output.push(output[output.len() - distance]);
                    }
                }
            }
        }
    }
    Ok(output)
}

fn offset_bits(size: usize) -> u32 {
    let mut bits = 4;
    while (1usize << bits) < size {
        bits += 1;
    }
    bits
}

/// Greedy, deterministic encoder. MS-OVBA 2.4.1.3.10 requires zero padding
/// when an incompressible final chunk has fewer than 4096 input bytes.
pub fn compress(input: &[u8]) -> Result<Vec<u8>> {
    let mut output = vec![1];
    for chunk in input.chunks(4096) {
        let mut encoded = Vec::new();
        let mut at = 0;
        while at < chunk.len() {
            let flag_at = encoded.len();
            encoded.push(0);
            for bit in 0..8 {
                if at == chunk.len() {
                    break;
                }
                let bits = offset_bits(at);
                let max_length = ((0xffffu16 >> bits) as usize + 3).min(chunk.len() - at);
                let mut best = (0, 0);
                for distance in 1..=at.min(1 << bits) {
                    if chunk[at - distance] != chunk[at] {
                        continue;
                    }
                    let mut length = 1;
                    while length < max_length && chunk[at + length - distance] == chunk[at + length]
                    {
                        length += 1;
                    }
                    if length >= 3 && length > best.1 {
                        best = (distance, length);
                        if length == max_length {
                            break;
                        }
                    }
                }
                if best.1 >= 3 {
                    encoded[flag_at] |= 1 << bit;
                    let token = (((best.0 - 1) as u16) << (16 - bits)) | ((best.1 - 3) as u16);
                    encoded.extend_from_slice(&token.to_le_bytes());
                    at += best.1;
                } else {
                    encoded.push(chunk[at]);
                    at += 1;
                }
            }
        }
        if encoded.len() <= 4096 {
            let header = 0xb000 | (encoded.len() as u16 - 1);
            output.extend_from_slice(&header.to_le_bytes());
            output.extend_from_slice(&encoded);
        } else {
            output.extend_from_slice(&0x3fffu16.to_le_bytes());
            output.extend_from_slice(chunk);
            output.resize(output.len() + 4096 - chunk.len(), 0);
        }
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn published_literal_and_copy_examples() {
        assert_eq!(
            decompress(&[1, 3, 0xb0, 0, b'a', b'b', b'c'], 20).unwrap(),
            b"abc"
        );
        assert_eq!(
            decompress(&[1, 5, 0xb0, 8, b'a', b'b', b'c', 3, 0x20], 20).unwrap(),
            b"abcabcabc"
        );
    }
    #[test]
    fn chunk_boundaries_and_overlap_round_trip() {
        for size in [0, 1, 15, 16, 17, 255, 256, 257, 4095, 4096, 4097, 20000] {
            let source = (0..size)
                .map(|n| b"Attribute VB_Name = \"Module1\"\r\n"[n % 31])
                .collect::<Vec<_>>();
            assert_eq!(
                decompress(&compress(&source).unwrap(), size).unwrap(),
                source
            );
        }
    }
    #[test]
    fn malformed_and_bounded_containers() {
        for bytes in [
            &[0][..],
            &[1, 0][..],
            &[1, 0, 0][..],
            &[1, 2, 0xb0, 1, 0, 0][..],
        ] {
            assert!(decompress(bytes, 4096).is_err());
        }
        let packed = compress(&[b'a'; 9000]).unwrap();
        assert!(decompress(&packed, 8999).is_err());
        for end in 2..packed.len() {
            let _ = decompress(&packed[..end], 9000);
        }
    }
    #[test]
    fn incompressible_final_raw_chunk_has_the_specified_zero_padding() {
        let mut state = 0x18273465u32;
        let data = (0..4000)
            .map(|_| {
                state ^= state << 13;
                state ^= state >> 17;
                state ^= state << 5;
                state as u8
            })
            .collect::<Vec<_>>();
        let compressed = compress(&data).unwrap();
        assert_eq!(&compressed[..3], &[1, 0xff, 0x3f]);
        let decoded = decompress(&compressed, 4096).unwrap();
        assert_eq!(&decoded[..data.len()], data);
        assert!(decoded[data.len()..].iter().all(|b| *b == 0));
    }
}
