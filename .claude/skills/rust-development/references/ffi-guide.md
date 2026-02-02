# FFI and C Interop Guide

## Calling C from Rust

### Basic FFI
```rust
// Link to system library
#[link(name = "c")]
extern "C" {
    fn strlen(s: *const c_char) -> usize;
    fn malloc(size: usize) -> *mut c_void;
    fn free(ptr: *mut c_void);
}

use std::ffi::{c_char, c_void, CString};

fn main() {
    let s = CString::new("hello").unwrap();
    let len = unsafe { strlen(s.as_ptr()) };
}
```

### Bindgen for Auto-Generation
```toml
# Cargo.toml
[build-dependencies]
bindgen = "0.69"
```

```rust
// build.rs
use std::path::PathBuf;

fn main() {
    println!("cargo:rerun-if-changed=wrapper.h");
    
    let bindings = bindgen::Builder::default()
        .header("wrapper.h")
        .parse_callbacks(Box::new(bindgen::CargoCallbacks::new()))
        .generate()
        .expect("Unable to generate bindings");
    
    let out_path = PathBuf::from(std::env::var("OUT_DIR").unwrap());
    bindings
        .write_to_file(out_path.join("bindings.rs"))
        .expect("Couldn't write bindings");
}
```

```rust
// src/lib.rs
include!(concat!(env!("OUT_DIR"), "/bindings.rs"));
```

## Exposing Rust to C

### C-Compatible Types
```rust
#[repr(C)]
pub struct Point {
    pub x: f64,
    pub y: f64,
}

#[repr(C)]
pub enum Status {
    Ok = 0,
    Error = 1,
}
```

### Exported Functions
```rust
#[no_mangle]
pub extern "C" fn add(a: i32, b: i32) -> i32 {
    a + b
}

#[no_mangle]
pub extern "C" fn process_point(p: *const Point) -> f64 {
    if p.is_null() {
        return 0.0;
    }
    unsafe {
        (*p).x + (*p).y
    }
}
```

### Library Configuration
```toml
# Cargo.toml
[lib]
crate-type = ["cdylib", "staticlib"]
```

### Header Generation with cbindgen
```toml
# cbindgen.toml
language = "C"
include_guard = "MY_LIB_H"

[export]
include = ["Point", "Status"]
```

```bash
cbindgen --config cbindgen.toml --output my_lib.h
```

## String Handling

### Rust to C
```rust
use std::ffi::CString;

fn rust_to_c() {
    let rust_str = "hello";
    let c_str = CString::new(rust_str).expect("CString failed");
    
    unsafe {
        some_c_function(c_str.as_ptr());
    }
}
```

### C to Rust
```rust
use std::ffi::{CStr, c_char};

unsafe fn c_to_rust(ptr: *const c_char) -> String {
    if ptr.is_null() {
        return String::new();
    }
    CStr::from_ptr(ptr)
        .to_str()
        .unwrap_or("")
        .to_owned()
}
```

### Returning Strings to C
```rust
#[no_mangle]
pub extern "C" fn get_message() -> *mut c_char {
    let s = CString::new("Hello from Rust").unwrap();
    s.into_raw()  // Caller must free
}

#[no_mangle]
pub extern "C" fn free_string(ptr: *mut c_char) {
    if !ptr.is_null() {
        unsafe { drop(CString::from_raw(ptr)); }
    }
}
```

## Memory Management

### Opaque Types
```rust
pub struct Context {
    data: Vec<u8>,
}

#[no_mangle]
pub extern "C" fn context_new() -> *mut Context {
    Box::into_raw(Box::new(Context { data: vec![] }))
}

#[no_mangle]
pub extern "C" fn context_free(ctx: *mut Context) {
    if !ctx.is_null() {
        unsafe { drop(Box::from_raw(ctx)); }
    }
}

#[no_mangle]
pub extern "C" fn context_add(ctx: *mut Context, value: u8) {
    if let Some(ctx) = unsafe { ctx.as_mut() } {
        ctx.data.push(value);
    }
}
```

## Error Handling

```rust
#[repr(C)]
pub struct Result {
    pub success: bool,
    pub error_code: i32,
    pub error_message: *mut c_char,
}

impl Result {
    fn ok() -> Self {
        Self { success: true, error_code: 0, error_message: std::ptr::null_mut() }
    }
    
    fn err(code: i32, msg: &str) -> Self {
        Self {
            success: false,
            error_code: code,
            error_message: CString::new(msg).unwrap().into_raw(),
        }
    }
}

#[no_mangle]
pub extern "C" fn do_operation() -> Result {
    match internal_operation() {
        Ok(_) => Result::ok(),
        Err(e) => Result::err(1, &e.to_string()),
    }
}
```

## Callbacks

```rust
type Callback = extern "C" fn(i32) -> i32;

#[no_mangle]
pub extern "C" fn register_callback(cb: Callback) {
    let result = cb(42);
    println!("Callback returned: {}", result);
}

// With user data
type CallbackWithData = extern "C" fn(*mut c_void, i32) -> i32;

#[no_mangle]
pub extern "C" fn register_callback_with_data(
    cb: CallbackWithData,
    user_data: *mut c_void,
) {
    let result = cb(user_data, 42);
}
```

## Platform-Specific Linking

```rust
#[cfg(target_os = "linux")]
#[link(name = "ssl")]
extern "C" { /* ... */ }

#[cfg(target_os = "macos")]
#[link(name = "Security", kind = "framework")]
extern "C" { /* ... */ }

#[cfg(target_os = "windows")]
#[link(name = "kernel32")]
extern "C" { /* ... */ }
```
