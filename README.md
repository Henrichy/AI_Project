
# AI Research Assistant MVP

An AI-powered research assistant that helps students from topic selection to final defense.

## Features (Phase 1)

1. **Topic Selection**: Generate 20 research topics with novelty scores, research gaps, and feasibility scores
2. **Proposal Generator**: Create complete research proposals with all necessary sections

## Getting Started

### Prerequisites

- Python 3.8+
- OpenAI API key

### Backend Setup

1. Navigate to the backend directory:
   ```bash
   cd backend
   ```

2. Activate the virtual environment:
   ```bash
   source venv/bin/activate  # On macOS/Linux
   # OR
   venv\Scripts\activate     # On Windows
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Create a `.env` file with your OpenAI API key:
   ```
   OPENAI_API_KEY=your_actual_api_key_here
   ```

5. Run the backend server:
   ```bash
   python main.py
   ```

The backend will be available at http://localhost:8000

### Frontend Setup

1. Open `frontend/index.html` in your browser, or serve it using a simple HTTP server:
   ```bash
   cd frontend
   python -m http.server 8080
   ```

The frontend will be available at http://localhost:8080

## Project Structure

```
projectai/
├── backend/
│   ├── main.py              # FastAPI application
│   ├── requirements.txt     # Python dependencies
│   ├── .env                 # Environment variables
│   └── venv/                # Virtual environment
├── frontend/
│   ├── index.html           # Main HTML file
│   ├── style.css            # Styles
│   └── script.js            # Frontend logic
└── README.md                # This file
```

## Future Phases

- **Phase 2**: RAG-based literature review assistant with PDF upload
- **Phase 3**: Data analysis, citation generation, and PowerPoint generation

## License

MIT
