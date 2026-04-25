# poi-decrypt-checker

Small Apache POI helper used by .NET interop tests and test vector generation.

## Build

```bash
mvn -q -DskipTests package
```

## Usage

```bash
JAR=target/poi-decrypt-checker-1.0.0-jar-with-dependencies.jar

# Decrypt
java -jar $JAR decrypt <encrypted.xlsx> <output.xlsx> <password>

# Encrypt with AES-256 / SHA-512 (agile)
java -jar $JAR encrypt <input.xlsx> <output.xlsx> <password>

# Create a plain XLSX (types: simple, formulas, styles, japanese)
java -jar $JAR create <type> <output.xlsx>
```

## Regenerating test vectors

Run from the repository root after building the JAR:

```bash
JAR=tests/poi-decrypt-checker/target/poi-decrypt-checker-1.0.0-jar-with-dependencies.jar

for type in simple formulas styles japanese; do
  java -jar "$JAR" create "$type" "test-vectors/plain/$type.xlsx"
  java -jar "$JAR" encrypt "test-vectors/plain/$type.xlsx" \
    "test-vectors/encrypted-by-apache-poi/${type}_aes256_sha512.xlsx" pass
done
```

Password used for all test vectors: `pass`
Encryption: AES-256 / SHA-512 (OOXML agile encryption)
