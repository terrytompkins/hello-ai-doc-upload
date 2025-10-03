import streamlit as st
import openai
import os
from dotenv import load_dotenv
import tempfile
import json
from pptx import Presentation

# Load environment variables
load_dotenv()

# Configure Streamlit page
st.set_page_config(
    page_title="AI Document Chat",
    page_icon="📄",
    layout="wide"
)

# Initialize session state
if "messages" not in st.session_state:
    st.session_state.messages = []
if "document_content" not in st.session_state:
    st.session_state.document_content = ""
if "api_key" not in st.session_state:
    st.session_state.api_key = ""

def get_openai_client(api_key):
    """Initialize OpenAI client with the provided API key"""
    if not api_key:
        return None
    try:
        client = openai.OpenAI(api_key=api_key)
        # Test the API key by making a simple request
        client.models.list()
        return client
    except Exception as e:
        st.error(f"Error connecting to OpenAI API: {str(e)}")
        return None

def process_document(file_content, filename):
    """Process the uploaded document and extract text content"""
    try:
        if filename.endswith('.txt') or filename.endswith('.md'):
            # Handle text and markdown files
            return file_content.decode('utf-8')
        elif filename.endswith('.pptx'):
            # Handle PowerPoint files
            return process_pptx_file(file_content)
        else:
            # For other file types, try to decode as text
            return file_content.decode('utf-8')
    except Exception as e:
        st.error(f"Error processing document: {str(e)}")
        return ""

def process_pptx_file(file_content):
    """Extract text content from a PowerPoint file"""
    try:
        # Create a temporary file to save the PowerPoint content
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as temp_file:
            temp_file.write(file_content)
            temp_file_path = temp_file.name
        
        # Load the presentation
        presentation = Presentation(temp_file_path)
        
        # Extract text from all slides
        extracted_text = []
        for i, slide in enumerate(presentation.slides, 1):
            slide_text = []
            slide_text.append(f"--- Slide {i} ---")
            
            # Extract text from all shapes in the slide
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_text.append(shape.text.strip())
            
            if len(slide_text) > 1:  # More than just the slide header
                extracted_text.extend(slide_text)
        
        # Clean up temporary file
        os.unlink(temp_file_path)
        
        return "\n".join(extracted_text) if extracted_text else "No text content found in the PowerPoint file."
        
    except Exception as e:
        st.error(f"Error processing PowerPoint file: {str(e)}")
        return ""

def chat_with_ai(client, user_message, document_context=""):
    """Send a message to OpenAI with optional document context"""
    try:
        system_prompt = "You are a helpful AI assistant. "
        if document_context:
            system_prompt += f"Here is some context from a document the user uploaded: {document_context}\n\n"
        system_prompt += "Please help the user with their questions about the document or any other topics."
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message}
            ],
            max_tokens=1000,
            temperature=0.7
        )
        
        return response.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"

# Main app layout
st.title("📄 AI Document Chat")
st.markdown("Upload a document and chat with AI about its content!")

# Create two columns
col1, col2 = st.columns([1, 2])

# Left column - Controls
with col1:
    st.header("Settings")
    
    # API Key input
    st.subheader("OpenAI API Key")
    api_key_input = st.text_input(
        "Enter your OpenAI API key:",
        value=st.session_state.api_key,
        type="password",
        help="You can also set the OPENAI_API_KEY environment variable"
    )
    
    # Update session state
    if api_key_input:
        st.session_state.api_key = api_key_input
    
    # Check for environment variable
    env_api_key = os.getenv("OPENAI_API_KEY")
    if env_api_key and not api_key_input:
        st.session_state.api_key = env_api_key
        st.info("Using API key from environment variable")
    
    # File upload
    st.subheader("Upload Document")
    uploaded_file = st.file_uploader(
        "Choose a document to upload:",
        type=['txt', 'md', 'pptx'],
        help="Supports .txt, .md, and .pptx files"
    )
    
    if uploaded_file is not None:
        # Read the file content
        file_content = uploaded_file.read()
        document_text = process_document(file_content, uploaded_file.name)
        
        if document_text:
            st.session_state.document_content = document_text
            st.success(f"✅ Document '{uploaded_file.name}' uploaded successfully!")
            st.text_area(
                "Document Preview:",
                value=document_text[:500] + "..." if len(document_text) > 500 else document_text,
                height=200,
                disabled=True
            )
        else:
            st.error("❌ Failed to process the document")
    
    # Clear document button
    if st.session_state.document_content:
        if st.button("🗑️ Clear Document"):
            st.session_state.document_content = ""
            st.session_state.messages = []
            st.rerun()

# Right column - Chat interface
with col2:
    st.header("AI Chat")
    
    # Check if API key is available
    if not st.session_state.api_key:
        st.warning("⚠️ Please enter your OpenAI API key in the left panel to start chatting.")
    else:
        # Initialize OpenAI client
        client = get_openai_client(st.session_state.api_key)
        
        if client is None:
            st.error("❌ Unable to connect to OpenAI API. Please check your API key.")
        else:
            # Display chat messages
            for message in st.session_state.messages:
                with st.chat_message(message["role"]):
                    st.markdown(message["content"])
            
            # Chat input
            if prompt := st.chat_input("Ask me anything about your document or any other topic!"):
                # Add user message to chat history
                st.session_state.messages.append({"role": "user", "content": prompt})
                
                # Display user message
                with st.chat_message("user"):
                    st.markdown(prompt)
                
                # Get AI response
                with st.chat_message("assistant"):
                    with st.spinner("Thinking..."):
                        response = chat_with_ai(
                            client, 
                            prompt, 
                            st.session_state.document_content
                        )
                    st.markdown(response)
                
                # Add assistant response to chat history
                st.session_state.messages.append({"role": "assistant", "content": response})

# Footer
st.markdown("---")
st.markdown("**Instructions:**")
st.markdown("1. Enter your OpenAI API key in the left panel")
st.markdown("2. Upload a document (supports .txt, .md, and .pptx files)")
st.markdown("3. Start chatting with the AI about your document!")
