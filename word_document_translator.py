import requests
import re
import html
import os
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

OLLAMA_URL = "http://localhost:11434/api/generate"
MODEL_CONTEXT = "qwen2.5:7b"

LANGUAGE_SOURCE = "english"
LANGUAGE_TARGET = "french"

REQUEST_TIMEOUT_SECONDS = 60

INPUT_DOCX_PATH = "original_document.docx"
OUTPUT_DOCX_PATH = "translated_document.docx"


# ---------------------------------------------------------------------------
# 1. OLLAMA COMMUNICATION FUNCTIONS
# ---------------------------------------------------------------------------

class OllamaUnavailableError(Exception):
    """Raised when the Ollama service cannot be reached."""


class OllamaQueryError(Exception):
    """Raised when a request to Ollama fails (timeout, HTTP error, etc.)."""


def check_ollama_reachable(ollama_url=OLLAMA_URL, timeout=5):
    """Checks if the Ollama service is reachable before translation starts."""
    # OLLAMA_URL usually ends with /api/generate; /api/tags is a lightweight health endpoint.
    base_url = ollama_url.replace("/api/generate", "")
    health_url = f"{base_url}/api/tags"
    try:
        response = requests.get(health_url, timeout=timeout)
        response.raise_for_status()
    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as exc:
        raise OllamaUnavailableError(
            f"Ollama is unreachable at {health_url}. "
            f"Check that the Ollama service is running and accessible."
        ) from exc
    except requests.exceptions.RequestException as exc:
        raise OllamaUnavailableError(
            f"Ollama health check failed ({health_url}): {exc}"
        ) from exc

def generate_document_context(full_text, ollama_url=OLLAMA_URL, model_context=MODEL_CONTEXT, source_language=LANGUAGE_SOURCE, target_language=LANGUAGE_TARGET):
    """Analyzes the entire document to create a universal context summary."""
    url = ollama_url
    prompt = (
        f"Analyze the following text extracted from a document. "
        f"Generate a brief summary (2-3 sentences max) describing only: "
        f"1. The main subject of the document.\n"
        f"2. The tone to adopt (formal, literary, technical, legal, marketing, etc.).\n"
        f"3. The target audience.\n"
        f"Be extremely concise and objective.\n\nDocument text:\n{full_text}"
    )
    payload = {
        "model": model_context,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.3
        }
    }
    try:
        response = requests.post(url, json=payload, timeout=REQUEST_TIMEOUT_SECONDS)
        response.raise_for_status()
        return response.json().get("response", "").strip()
    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
        print(f"Error during context generation (Ollama query failed): {e}")
        raise OllamaQueryError("Context generation failed because Ollama query did not complete.") from e
    except requests.exceptions.RequestException as e:
        print(f"Error during context generation (HTTP/API failure): {e}")
        raise OllamaQueryError("Context generation failed due to an Ollama API error.") from e
    except ValueError as e:
        print(f"Error during context generation (invalid JSON response): {e}")
        return "No specific context."

def translate_text_with_context(tagged_text, context, ollama_url=OLLAMA_URL, model_context=MODEL_CONTEXT, source_language=LANGUAGE_SOURCE, target_language=LANGUAGE_TARGET):
    """Translates the text generically while filtering out AI conversational preambles."""
    # Anti-hallucination security: avoid calling the API if the text has no letters
    if not re.search(r'[a-zA-Z]', tagged_text):
        return tagged_text

    url = ollama_url
    
    # Restructured prompt to "force" direct translation start
    prompt = f"""You are an expert translator.
    Global context of the document: "{context}"

    ABSOLUTE AND IMPERATIVE RULES:
    1. Translate the ENTIRE text FROM {source_language} into {target_language}. Leave absolutely NO word, part of a sentence, or start of a sentence in {source_language}. EVERYTHING must be translated.
    2. Pay close attention to quotes (" or '): you must translate everything before, inside, and after the quotes.
    3. Keep EXACTLY the position of HTML tags (<...>) without translating their code.
    4. NEVER copy the source text in {source_language} into your response.
    5. FORMAL PROHIBITION to add introductory or polite words.

    Original text:
    {tagged_text}

    Direct {target_language} translation:"""

    payload = {
        "model": model_context,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.3
        }
    }
    
    try:
        response = requests.post(url, json=payload, timeout=REQUEST_TIMEOUT_SECONDS)
        response.raise_for_status()
        translated_text = response.json().get("response", "").strip()
        
        # --- POST-TRANSLATION CLEANUP (Anti-preamble) ---
        # Remove typical introductory sentences if the model disobeys
        patterns_to_remove = [
            # French patterns
            r"^(Voici|Voilà) la traduction(.*?):\s*",
            r"^Voici le texte traduit(.*?):\s*",
            r"^Traduction(.*?):\s*",
            r"^Bien sûr(.*?):\s*",
            r"^Je suis désolé(.*?)\n",
            # English patterns
            r"^(Here is|Here's) the translation(.*?):\s*",
            r"^(Here is|Here's) the translated text(.*?):\s*",
            r"^Translation(.*?):\s*",
            r"^(Sure|Of course)(.*?):\s*",
            r"^I am sorry(.*?)\n"
        ]
        
        for pattern in patterns_to_remove:
            translated_text = re.sub(pattern, "", translated_text, flags=re.IGNORECASE | re.DOTALL).strip()
            
        return translated_text

    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
        print(f"Error during translation (Ollama query failed): {e}")
        return tagged_text
    except requests.exceptions.RequestException as e:
        print(f"Error during translation (HTTP/API failure): {e}")
        return tagged_text
    except ValueError as e:
        print(f"Error during translation (invalid JSON response): {e}")
        return tagged_text

def free_ollama_memory(ollama_url=OLLAMA_URL, model=MODEL_CONTEXT):
    """Forces Ollama to unload the model from memory (VRAM/RAM)."""
    payload = {
        "model": model,
        "keep_alive": 0
    }
    try:
        requests.post(ollama_url, json=payload)
        print(f" -> Memory successfully freed for the model {model}.")
    except Exception as e:
        print(f" -> Error freeing memory: {e}")

# ---------------------------------------------------------------------------
# 2. TAG MANIPULATION AND PYTHON-DOCX LOGIC
# ---------------------------------------------------------------------------

def paragraph_to_tags(paragraph):
    """Transforms a paragraph of runs into a tagged text string (HTML-like)."""
    tagged_text = ""
    for run in paragraph.runs:
        text = run.text
        if not text:
            continue
        # Escape potential existing tags to avoid XML conflicts
        text = text.replace("<", "&lt;").replace(">", "&gt;")
        
        # Apply tags according to the run's style
        if run.bold: text = f"<b>{text}</b>"
        if run.italic: text = f"<i>{text}</i>"
        if run.underline: text = f"<u>{text}</u>"
        if run.font.strike: text = f"<s>{text}</s>"
        
        tagged_text += text
    return tagged_text

def apply_tagged_translation(paragraph, translated_text):
    """Parses the translated text, filters invalid tags, and cleanly rebuilds the runs."""
    # Clean up potential markdown blocks hallucinated by the LLM
    translated_text = re.sub(r'```[a-zA-Z0-9]*', '', translated_text).strip()
    
    # Empty existing runs
    p_element = paragraph._p
    for run in paragraph.runs:
        p_element.remove(run._r)
        
    # Global regex that captures any HTML tag
    tag_pattern = r'(</?\w+(?:\s+[^>]+)?>)'
    tokens = re.split(tag_pattern, translated_text)
    
    bold_level = 0
    italic_level = 0
    underline_level = 0
    strike_level = 0
    
    for token in tokens:
        if not token:
            continue
            
        # If the token is an HTML tag
        if token.startswith("<") and token.endswith(">"):
            match_name = re.match(r'</?(\w+)', token)
            if not match_name:
                continue
            tag_name = match_name.group(1).lower()
            is_closing = token.startswith("</")
            
            # Handle styles and common aliases (strong/bold, em/italic)
            if tag_name in ["b", "strong"]:
                bold_level += -1 if is_closing else 1
            elif tag_name in ["i", "em"]:
                italic_level += -1 if is_closing else 1
            elif tag_name in ["u"]:
                underline_level += -1 if is_closing else 1
            elif tag_name in ["s", "strike", "del"]:
                strike_level += -1 if is_closing else 1
            
            # Safeguard to prevent going below 0
            bold_level = max(0, bold_level)
            italic_level = max(0, italic_level)
            underline_level = max(0, underline_level)
            strike_level = max(0, strike_level)
            
            # Other tags (<span>, <div>, etc.) are silently ignored.
            
        else:
            # It's raw text: convert HTML entities (e.g., &amp; -> &)
            clean_text = html.unescape(token)
            
            if clean_text:
                run = paragraph.add_run(clean_text)
                if bold_level > 0: run.bold = True
                if italic_level > 0: run.italic = True
                if underline_level > 0: run.underline = True
                if strike_level > 0: run.font.strike = True

# ---------------------------------------------------------------------------
# 3. MAIN PROCESS
# ---------------------------------------------------------------------------

def extract_all_text(input_path):
    """Scans the document a first time to extract all raw text."""
    doc = Document(input_path)
    texts = []
    # Limit the extracted size to avoid overloading AI model's context window
    for p in doc.paragraphs:
        if p.text.strip(): texts.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text.strip(): texts.append(p.text)
    for txbx in doc.element.xpath('//w:txbxContent'):
        for p_element in txbx.xpath('.//w:p'):
            p = Paragraph(p_element, doc._body)
            if p.text.strip(): texts.append(p.text)
            
    # Returns the first 8000 characters to give a general idea without crashing the API
    return "\n".join(texts)[:8000]

def translate_full_docx(input_path, output_path, ollama_url, model_context, source_language, target_language):
    if not os.path.isfile(input_path):
        print(f"Error: input file not found: {input_path}")
        return False

    try:
        check_ollama_reachable(ollama_url=ollama_url)
    except OllamaUnavailableError as e:
        print(f"Error: {e}")
        return False

    print("1. Extracting text for context analysis...")
    try:
        full_text = extract_all_text(input_path)
    except FileNotFoundError:
        print(f"Error: input file not found: {input_path}")
        return False
    except Exception as e:
        print(f"Error while reading input document: {e}")
        return False
    
    print(f"2. Generating context with AI model ({model_context})...")
    try:
        context = generate_document_context(full_text, ollama_url=ollama_url, model_context=model_context, source_language=source_language, target_language=target_language)
    except OllamaQueryError as e:
        print(f"Warning: {e}")
        context = "No specific context."
    print(f"   [Retained context] : {context}\n")
    
    try:
        doc = Document(input_path)
    except FileNotFoundError:
        print(f"Error: input file not found: {input_path}")
        return False
    except Exception as e:
        print(f"Error while opening input document: {e}")
        return False
    
    print("3. Translating paragraphs...")
    for paragraph in doc.paragraphs:
        tagged_text = paragraph_to_tags(paragraph)
        if tagged_text.strip() and re.search(r'[a-zA-Z]', tagged_text):
            translation = translate_text_with_context(tagged_text, context, ollama_url=ollama_url, model_context=model_context, source_language=source_language, target_language=target_language)
            apply_tagged_translation(paragraph, translation)
            
    print("4. Translating tables...")
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    tagged_text = paragraph_to_tags(paragraph)
                    if tagged_text.strip() and re.search(r'[a-zA-Z]', tagged_text):
                        translation = translate_text_with_context(tagged_text, context, ollama_url=ollama_url, model_context=model_context, source_language=source_language, target_language=target_language)
                        apply_tagged_translation(paragraph, translation)

    print("5. Translating text boxes...")
    for txbx in doc.element.xpath('//w:txbxContent'):
        for p_element in txbx.xpath('.//w:p'):
            paragraph = Paragraph(p_element, doc._body)
            tagged_text = paragraph_to_tags(paragraph)
            if tagged_text.strip() and re.search(r'[a-zA-Z]', tagged_text):
                translation = translate_text_with_context(tagged_text, context, ollama_url=ollama_url, model_context=model_context, source_language=source_language, target_language=target_language)
                apply_tagged_translation(paragraph, translation)

    print("6. Saving the translated document...")
    try:
        doc.save(output_path)
    except PermissionError as e:
        print(f"Error: translated output file cannot be generated (permission denied): {output_path}")
        print(f"Details: {e}")
        return False
    except OSError as e:
        print(f"Error: translated output file cannot be generated: {output_path}")
        print(f"Details: {e}")
        return False

    print(f"Translation complete! File saved as: {output_path}")

    print("7. Unloading the model from memory...")
    free_ollama_memory(ollama_url=ollama_url, model=model_context)
    return True

if __name__ == "__main__":
    success = translate_full_docx(INPUT_DOCX_PATH, OUTPUT_DOCX_PATH, OLLAMA_URL, MODEL_CONTEXT, LANGUAGE_SOURCE, LANGUAGE_TARGET)
    if not success:
        print("Translation aborted due to one or more errors.")