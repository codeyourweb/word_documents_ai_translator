# Microsoft Word documents AI translator

This project is an experimentation to translate large Word documents (`.docx`) while preserving style and structure, using a local AI model through Ollama.

It is designed to keep formatting elements such as bold, italic, underline, strikethrough, tables, and text boxes as intact as possible during translation.

## Why this project

Translating long `.docx` files often breaks formatting or requires cloud services.
This script focuses on:

- Processing large Word documents locally.
- Preserving document structure and styling.
- Using a local LLM workflow (no mandatory external translation API).

## Current defaults

Default values (all overridable via CLI arguments):

| Argument | Default |
|---|---|
| `--source-lang` | `french` |
| `--target-lang` | `english` |
| `--input` | `original_document.docx` |
| `--output` | `translated_document.docx` |
| `--ollama-url` | `http://localhost:11434` |
| `--model` | `qwen2.5:7b` |
| `--timeout` | `120` (seconds) |

Source and target language support depends only on the AI model used.

## Requirements

- Python 3.9+
- [Ollama](https://ollama.com/) running locally
- A pulled model in Ollama (default: `qwen2.5:7b`)

Install Python dependencies:

```bash
pip install -r requirements.txt
```

Pull default model (if not already available):

```bash
ollama pull qwen2.5:7b
```

The default model is **`qwen2.5:7b`** (Qwen 2.5), selected as the best balance between power and lightness for this use case.

## Usage

```bash
python word_document_translator.py
```

Or override any option via CLI arguments:

```bash
python word_document_translator.py \
  --input my_document.docx \
  --output translated.docx \
  --source-lang english \
  --target-lang spanish \
  --model qwen2.5:7b \
  --ollama-url http://localhost:11434 \
  --timeout 120
```

### CLI arguments

| Argument | Description | Default |
|---|---|---|
| `--input FILE` | Input `.docx` file path | `original_document.docx` |
| `--output FILE` | Output `.docx` file path | `translated_document.docx` |
| `--source-lang LANG` | Source language | `english` |
| `--target-lang LANG` | Target language | `french` |
| `--model NAME` | Ollama model name | `qwen2.5:7b` |
| `--ollama-url URL` | Ollama base URL | `http://localhost:11434` |
| `--timeout SECONDS` | Request timeout in seconds | `120` |

On success, the translated file is saved to the output path.

## Notes and limitations

- This is an experimental project and translation quality depends on model behavior.
- Very complex formatting can still require manual review after translation.
- Performance and quality vary with document length and model hardware constraints.

## Maintenance policy

While this code is shared publicly, it is an in-house development for personal use and will not necessarily be maintained.

Pull requests are welcome to help improve this project.

# License
This project is licensed under the GNU AFFERO GENERAL PUBLIC LICENSE. See LICENSE for details.