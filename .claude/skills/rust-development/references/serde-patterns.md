# Serde Patterns

## Basic Derive
```rust
use serde::{Deserialize, Serialize};

#[derive(Serialize, Deserialize)]
struct User {
    id: i64,
    name: String,
    #[serde(default)]
    active: bool,
}
```

## Field Attributes

```rust
#[derive(Serialize, Deserialize)]
struct Config {
    #[serde(rename = "apiKey")]
    api_key: String,
    
    #[serde(skip_serializing_if = "Option::is_none")]
    optional_field: Option<String>,
    
    #[serde(default = "default_timeout")]
    timeout: u64,
    
    #[serde(flatten)]
    extra: HashMap<String, Value>,
    
    #[serde(skip)]
    internal: InternalState,
}

fn default_timeout() -> u64 { 30 }
```

## Enum Representations

```rust
// Externally tagged (default)
#[derive(Serialize, Deserialize)]
enum Message {
    Request { id: u64 },
    Response { data: String },
}
// {"Request": {"id": 1}}

// Internally tagged
#[derive(Serialize, Deserialize)]
#[serde(tag = "type")]
enum Event {
    Click { x: i32, y: i32 },
    KeyPress { key: String },
}
// {"type": "Click", "x": 10, "y": 20}

// Adjacently tagged
#[derive(Serialize, Deserialize)]
#[serde(tag = "t", content = "c")]
enum Payload {
    Text(String),
    Binary(Vec<u8>),
}
// {"t": "Text", "c": "hello"}

// Untagged
#[derive(Serialize, Deserialize)]
#[serde(untagged)]
enum StringOrInt {
    Int(i64),
    String(String),
}
```

## Custom Serialization

```rust
use serde::{Serializer, Deserializer};

#[derive(Serialize, Deserialize)]
struct Record {
    #[serde(serialize_with = "serialize_date")]
    #[serde(deserialize_with = "deserialize_date")]
    date: NaiveDate,
}

fn serialize_date<S>(date: &NaiveDate, s: S) -> Result<S::Ok, S::Error>
where S: Serializer {
    s.serialize_str(&date.format("%Y-%m-%d").to_string())
}

fn deserialize_date<'de, D>(d: D) -> Result<NaiveDate, D::Error>
where D: Deserializer<'de> {
    let s = String::deserialize(d)?;
    NaiveDate::parse_from_str(&s, "%Y-%m-%d")
        .map_err(serde::de::Error::custom)
}
```

## Visitor Pattern for Complex Deserialization

```rust
use serde::de::{self, Visitor, MapAccess};

struct MyVisitor;

impl<'de> Visitor<'de> for MyVisitor {
    type Value = MyType;
    
    fn expecting(&self, f: &mut fmt::Formatter) -> fmt::Result {
        f.write_str("a map with specific fields")
    }
    
    fn visit_map<M>(self, mut map: M) -> Result<Self::Value, M::Error>
    where M: MapAccess<'de> {
        let mut id = None;
        while let Some(key) = map.next_key::<String>()? {
            match key.as_str() {
                "id" => id = Some(map.next_value()?),
                _ => { let _ = map.next_value::<de::IgnoredAny>()?; }
            }
        }
        Ok(MyType { id: id.ok_or_else(|| de::Error::missing_field("id"))? })
    }
}
```

## JSON/YAML/TOML

```rust
// JSON
let json = serde_json::to_string_pretty(&data)?;
let parsed: Data = serde_json::from_str(&json)?;

// YAML
let yaml = serde_yaml::to_string(&data)?;
let parsed: Data = serde_yaml::from_str(&yaml)?;

// TOML
let toml = toml::to_string_pretty(&data)?;
let parsed: Data = toml::from_str(&toml)?;
```
