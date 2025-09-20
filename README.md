# Visual Novel Companion (VNC)

Learn your target language by reading visual novels. Translate dialog, extract vocabulary, export it to Anki or review it in-app with context.

## Translation

<img width="2662" height="1443" alt="image" src="https://github.com/user-attachments/assets/ce59a474-121d-45d2-8985-25cdc2757d76" />

## Export vocabulary lists to Anki or CSV.

<img width="625" height="643" alt="image" src="https://github.com/user-attachments/assets/27056b45-5d42-4e8e-96a2-624fbfae29a6" />

## Review words in context

<img width="835" height="451" alt="image" src="https://github.com/user-attachments/assets/9b486698-43ef-4b6e-9075-08eb610de101" />

## Installation

Download the latest executable from the [releases page](https://github.com/cesar-bravo-m/VisualNovelCompanion/releases/tag/major) or [itch.io](https://cesar-bravo-m.itch.io/visual-novel-companion)

## Settings

The app needs an LLM to run, for which there are three options (selectable in-app):

- **Managed**: Use the LLM service I've exposed for free. It only has around $10/mo credits, which are shared among all users.
- **BYOK**: Use your own Together.ai API key
- **Local**: Use a local LLM running on ollama

*BYOK* and *Local* modes allow you to choose between two input modalities: Images or OCR.

Images is much more accurate, but uses more tokens and might be slower.

OCR uses Windows' built-in OCR model and sends the result for processing to the LLM. It uses far less tokens, but can be inaccurate.

## Usage

1) Open a visual novel
2) Open VNC, and choose the window of your visual novel
3) Click on "Translate" to translate the text and extract vocabulary

## Caveats

There is no vocabulary persistence yet. Export your lists to CSV or Anki after every play session.
