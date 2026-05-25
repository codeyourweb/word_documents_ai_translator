# Microsoft Word documents AI translator

This project is an experimentation to translate large Word documents (`.docx`) while preserving style and structure, using a local AI model through Ollama.

It is designed to keep formatting elements such as bold, italic, underline, strikethrough, tables, and text boxes as intact as possible during translation.

## Why this project

Translating long `.docx` files often breaks formatting or requires cloud services.
This script focuses on:

- Processing large Word documents locally.
- Preserving document structure and styling.
- Using a local LLM workflow (no mandatory external translation API).

## Default model

The default model is **`qwen2.5:7b`** (Qwen 2.5), selected as the best balance between power and lightness for this use case.

You can change the model in `word_document_translator.py` by editing `MODEL_CONTEXT` constant

## Current defaults

In `word_document_translator.py`, default values are:

- Source language: `english`
- Target language: `french`
- Input file: `original_document.docx`
- Output file: `translated_document.docx`
- Ollama endpoint: `http://localhost:11434`

Replace the variables with your desired values in the script. Source and target language support depends only on the AI model used. 

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

## Usage

1. Place your input document in the project root (or update `INPUT_DOCX_PATH`).
2. Ensure Ollama is running locally.
3. Run the translator:

```bash
python word_document_translator.py
```

If successful, the translated document will be written to `translated_document.docx` (or your configured output path).

## Notes and limitations

- This is an experimental project and translation quality depends on model behavior.
- Very complex formatting can still require manual review after translation.
- Performance and quality vary with document length and model hardware constraints.

## Maintenance policy

While this code is shared publicly, it is an in-house development for personal use and will not necessarily be maintained.

Pull requests are welcome to help improve this project.

# License
This project is licensed under the GNU AFFERO GENERAL PUBLIC LICENSE. See LICENSE for details.