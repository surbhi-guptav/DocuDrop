import os
import base64
import re
from openai import OpenAI
from dotenv import load_dotenv
import time

# Load API key from .env
load_dotenv()
api_key = os.getenv("OPENROUTER_API_KEY")
CAPTION_RATE_LIMIT_UNTIL = 0

if not api_key:
    print("❌ ERROR: API key not found. Add it inside .env file.")
    exit()

# Initialize client
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=api_key
)

# --- helpers to build data URLs ---
def _encode_image_path(image_path: str, mime_type: str = "image/png") -> str:
    with open(image_path, "rb") as img:
        encoded = base64.b64encode(img.read()).decode("utf-8")
    return f"data:{mime_type};base64,{encoded}"

def _encode_image_bytes(image_bytes: bytes, mime_type: str = "image/png") -> str:
    encoded = base64.b64encode(image_bytes).decode("utf-8")
    return f"data:{mime_type};base64,{encoded}"


def _content_to_text(content) -> str:
    """Normalize OpenAI/OpenRouter content into a plain string."""
    if isinstance(content, str):
        return content

    if isinstance(content, list):
        parts = []
        for part in content:
            if hasattr(part, "text") and isinstance(part.text, str):
                parts.append(part.text)
            elif isinstance(part, dict) and "text" in part:
                parts.append(part["text"])
        if parts:
            return "".join(parts)
        return " ".join(str(p) for p in content)

    return "" if content is None else str(content)


def _clean_caption(text: str) -> str:
    """Extract a single, clean, short caption from the model output."""
    if not text:
        return ""

    text = text.strip()
    if not text:
        return ""

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not lines:
        return ""

    first_line = lines[0]

    # Remove bullets / numbering / "Option 1:" style prefixes
    first_line = re.sub(r"^(option\s*\d+[:.)-]\s*)", "", first_line, flags=re.IGNORECASE)
    first_line = re.sub(r"^[\-\*\d\.\)\s]+", "", first_line)

    # Remove surrounding quotes if any
    first_line = first_line.strip(" '\"")

    # Enforce short caption: at most 7 words
    words = first_line.split()
    if len(words) > 7:
        first_line = " ".join(words[:7])

    return first_line


def _caption_from_data_url(image_data_url: str) -> str:
    """Core logic: call the model with a prepared data URL and return caption (or '')."""
    global CAPTION_RATE_LIMIT_UNTIL

    # If we already know the API is rate-limited, skip calling it
    if time.time() < CAPTION_RATE_LIMIT_UNTIL:
        remaining = int(CAPTION_RATE_LIMIT_UNTIL - time.time())
        print(f"[caption] Rate limited. Retry in {remaining}s.")
        return ""

    max_attempts = 3
    for attempt in range(1, max_attempts + 1):
        print(f"⏳ Sending image to model... (attempt {attempt}/{max_attempts})")

        try:
            response = client.chat.completions.create(
                model="google/gemma-3-27b-it:free",
                temperature=0.2,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": (
                                    "You are generating short, professional captions for technical "
                                    "screenshots in academic lab reports.\n\n"
                                    "Produce exactly ONE caption.\n"
                                    "- The caption must be a short noun phrase (not a full sentence).\n"
                                    "- Length: 3 to 7 words.\n"
                                    "- No bullets, numbering, labels, or explanations.\n"
                                    "- Do NOT write 'Option', 'Caption', or similar words.\n"
                                    "- Do NOT use markdown, emojis, or hashtags.\n"
                                    "Examples of the desired style:\n"
                                    "- Security group inbound rules configuration\n"
                                    "- AWS security group inbound rule set\n"
                                    "Now generate the caption for this image."
                                ),
                            },
                            {
                                "type": "image_url",
                                "image_url": {"url": image_data_url},
                            },
                        ],
                    }
                ],
                extra_body={
                    "provider": {
                        "only": ["google-ai-studio"],  # or "modelrun"
                        "allow_fallbacks": False,
                    }
                },
            )
        except Exception as e:
            msg = str(e)

            # 🔴 FIX 1: detect ANY 429 and STOP immediately
            if "429" in msg:
                CAPTION_RATE_LIMIT_UNTIL = time.time() + 50
                print("[caption] Rate limit hit (429).")
                return ""   # 🚫 NO RETRIES

            # other errors → retry
            print(f"⚠️ API error on attempt {attempt}: {e}")
            time.sleep(1.5)
            continue


        raw_content = response.choices[0].message.content
        text = _content_to_text(raw_content)
        caption = _clean_caption(text)

        # Validate caption: ensure 3–7 words
        if caption and 3 <= len(caption.split()) <= 7:
            return caption

        print(f"⚠️ Invalid or empty caption on attempt {attempt}. Raw content: {raw_content}")
        time.sleep(1.5)

    return ""


# --- public functions you will import in app.py ---

def caption_image(image_path: str, mime_type: str = "image/png") -> str:
    """Caption an image from a file path."""
    data_url = _encode_image_path(image_path, mime_type=mime_type)
    return _caption_from_data_url(data_url)

def caption_image_from_bytes(image_bytes: bytes, mime_type: str = "image/png") -> str:
    """Caption an image from raw bytes (for clipboard screenshots)."""
    data_url = _encode_image_bytes(image_bytes, mime_type=mime_type)
    return _caption_from_data_url(data_url)


if __name__ == "__main__":
    local_image = "sample_image.png"
    caption = caption_image(local_image)
    print("\n📌 Caption Result:\n")
    print(caption if caption else "[No caption returned]")


















# Works with the Google gemini model only via OpenRouter

# import os
# import base64
# import re
# from openai import OpenAI
# from dotenv import load_dotenv
# import time

# # Load API key from .env
# load_dotenv()
# api_key = os.getenv("OPENROUTER_API_KEY")

# if not api_key:
#     print("❌ ERROR: API key not found. Add it inside .env file.")
#     exit()

# # Initialize client
# client = OpenAI(
#     base_url="https://openrouter.ai/api/v1",
#     api_key=api_key
# )

# def encode_image(image_path: str, mime_type="image/png"):
#     """Reads the local image and encodes to Base64 Data URL format."""
#     with open(image_path, "rb") as img:
#         encoded = base64.b64encode(img.read()).decode("utf-8")
#     return f"data:{mime_type};base64,{encoded}"

# def _content_to_text(content) -> str:
#     """Normalize OpenAI/OpenRouter content into a plain string."""
#     if isinstance(content, str):
#         return content

#     if isinstance(content, list):
#         parts = []
#         for part in content:
#             if hasattr(part, "text") and isinstance(part.text, str):
#                 parts.append(part.text)
#             elif isinstance(part, dict) and "text" in part:
#                 parts.append(part["text"])
#         # If we collected text parts, use them
#         if parts:
#             return "".join(parts)
#         # Fallback: stringify all parts
#         return " ".join(str(p) for p in content)

#     return "" if content is None else str(content)

# def _clean_caption(text: str) -> str:
#     """Extract a single, clean, short caption from the model output."""
#     if not text:
#         return ""

#     text = text.strip()
#     if not text:
#         return ""

#     # Use the first non-empty line
#     lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
#     if not lines:
#         return ""

#     first_line = lines[0]

#     # Remove bullets / numbering / "Option 1:" style prefixes
#     first_line = re.sub(r"^(option\s*\d+[:.)-]\s*)", "", first_line, flags=re.IGNORECASE)
#     first_line = re.sub(r"^[\-\*\d\.\)\s]+", "", first_line)

#     # Remove surrounding quotes if any
#     first_line = first_line.strip(" '\"")

#     # Enforce short caption: at most 7 words
#     words = first_line.split()
#     if len(words) > 7:
#         first_line = " ".join(words[:7])

#     return first_line

# def caption_image(image_path: str) -> str:
#     image_data_url = encode_image(image_path)

#     max_attempts = 3
#     for attempt in range(1, max_attempts + 1):
#         print(f"⏳ Sending image to model... (attempt {attempt}/{max_attempts})")

#         try:
#             response = client.chat.completions.create(
#     model="google/gemma-3-27b-it:free",
#     temperature=0.2,
#     messages=[
#         {
#             "role": "user",
#             "content": [
#                 {
#                     "type": "text",
#                     "text": (
#                         "You are generating short, professional captions for technical "
#                         "screenshots in academic lab reports.\n\n"
#                         "Produce exactly ONE caption.\n"
#                         "- The caption must be a short noun phrase (not a full sentence).\n"
#                         "- Length: 3 to 7 words.\n"
#                         "- No bullets, numbering, labels, or explanations.\n"
#                         "- Do NOT write 'Option', 'Caption', or similar words.\n"
#                         "- Do NOT use markdown, emojis, or hashtags.\n"
#                         "Examples of the desired style:\n"
#                         "- Security group inbound rules configuration\n"
#                         "- AWS security group inbound rule set\n"
#                         "Now generate the caption for this image."
#                     ),
#                 },
#                 {
#                     "type": "image_url",
#                     "image_url": {"url": image_data_url},
#                 },
#             ],
#         }
#     ],
#     # 👇 Force a specific provider here
#     extra_body={
#         "provider": {
#             "only": ["google-ai-studio"],      # <--- replace with the exact slug you copied
#             "allow_fallbacks": False,  # don't fall back to Google AI Studio etc.
#         }
#     },
# )

#         except Exception as e:
#             print(f"⚠️ API error on attempt {attempt}: {e}")
#             time.sleep(1.5)
#             continue

#         raw_content = response.choices[0].message.content
#         text = _content_to_text(raw_content)
#         caption = _clean_caption(text)

#         # Validate caption: ensure 3–7 words
#         if caption and 3 <= len(caption.split()) <= 7:
#             return caption

#         print(f"⚠️ Invalid or empty caption on attempt {attempt}. Raw content: {raw_content}")
#         time.sleep(1.5)

#     # After all attempts failed
#     return ""

# if __name__ == "__main__":
#     local_image = "sample_image.png"

#     caption = caption_image(local_image)

#     print("\n📌 Caption Result:\n")
#     print(caption if caption else "[No caption returned]")