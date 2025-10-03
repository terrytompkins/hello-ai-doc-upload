# AI Document Chat

A Streamlit-based web application that allows you to upload documents and chat with AI about their content using the OpenAI API.

## Features

- 📄 Document upload (supports .txt, .md, and .pptx files)
- 🤖 Interactive AI chat interface
- 🔑 API key management (environment variable or manual input)
- 💬 Context-aware conversations about uploaded documents

## Setup

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Set up your OpenAI API key:**
   
   **Option A: Environment variable (recommended)**
   ```bash
   export OPENAI_API_KEY="your_api_key_here"
   ```
   
   **Option B: Enter in the app**
   - Run the app and enter your API key in the left panel

3. **Run the application:**
   ```bash
   streamlit run app.py
   ```

## Usage

1. Open your browser to the Streamlit app (usually `http://localhost:8501`)
2. Enter your OpenAI API key in the left panel
3. Upload a document using the file uploader
4. Start chatting with the AI about your document!

## File Structure

```
hello-ai-doc-upload/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
└── README.md          # This file
```

## Requirements

- Python 3.7+
- OpenAI API key
- Internet connection for API calls

## Notes

- Currently supports text files (.txt), Markdown files (.md), and PowerPoint files (.pptx)
- The app stores chat history in session state
- Document content is included as context in AI conversations
- API key is stored in session state for the duration of the app session