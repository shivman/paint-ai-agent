import os
from dotenv import load_dotenv
import google.generativeai as genai

def test_gemini_api():
    # Load environment variables
    load_dotenv()
    api_key = os.getenv('GEMINI_API_KEY')
    
    if not api_key:
        print("❌ Error: GEMINI_API_KEY not found in .env file")
        return False
    
    print(f"API Key found: {api_key[:4]}...{api_key[-4:]}")
    
    try:
        # Configure the library
        genai.configure(api_key=api_key)
        
        # List available models
        print("\nListing available models:")
        for m in genai.list_models():
            print(f"- {m.name}")
        
        # Get the model
        model = genai.GenerativeModel('gemini-2.0-flash')
        
        # Test generation
        print("\nTesting text generation:")
        response = model.generate_content("Say 'Hello, testing Gemini API!'")
        print(f"Response: {response.text}")
        
        print("\n✅ Gemini API test successful!")
        return True
        
    except Exception as e:
        print(f"\n❌ Error testing Gemini API: {str(e)}")
        return False

if __name__ == "__main__":
    print("🚀 Starting Gemini API test...")
    test_gemini_api() 