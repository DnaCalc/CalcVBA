#[unsafe(no_mangle)]
pub extern "system" fn CurrencyValue() -> i64 { 1_234_567_890_123 }
#[unsafe(no_mangle)]
pub extern "system" fn DateValue() -> f64 { 45123.75 }
#[unsafe(no_mangle)]
pub extern "system" fn ByteValue() -> u8 { 251 }
#[unsafe(no_mangle)]
pub extern "system" fn BoolOne() -> i16 { 1 }
#[unsafe(no_mangle)]
pub extern "system" fn Mixed(a:i32,b:f64,c:f32,d:i64,e:i16,f:f64,g:i64,h:i16,i:u8,j:i64) -> i32 {
    if a!=19 { return -1001 } if b!=23.5 { return -1002 } if c!=12.5 { return -1003 }
    if d!=12_345_000 { return -1004 } if e!= -2 { return -1005 } if f!=45123.75 { return -1006 }
    if g!=4_294_967_338 { return -1007 } if h!= -1 { return -1008 } if i!=251 { return -1009 }
    if j!= -9_876_543_210_987 { return -1010 } 42
}
