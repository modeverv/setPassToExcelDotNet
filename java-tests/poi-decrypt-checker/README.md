# poi-decrypt-checker

Small Apache POI helper used by .NET interop tests.

## Build

```bash
mvn -q -DskipTests package
```

## Usage

```bash
java -jar target/poi-decrypt-checker-1.0.0-jar-with-dependencies.jar decrypt <encrypted.xlsx> <output.xlsx> <password>
```

