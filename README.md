# Ledgerly

A dependency-free frontend interface for a grounded financial data assistant using uploaded Excel workbooks.

The UI is designed for a backend that:

- imports Excel workbook sheets into a server-side tabular representation;
- generates read-only SQL against imported sheet rows;
- executes SQL and returns rows as evidence;
- asks `openai/gpt-oss-120b` through Groq to explain only those returned values;
- stores a summary of the current session and injects only that summary into the next prompt;
- keeps previous sessions out of the active context.

The visible sample answer is demo data and is labeled as ledger evidence. Replace the static response in `app.js` with your workbook-import and Groq response once the backend is available.

## Run

## Run the AI version

1. Install the Python dependency: `pip install -r requirements.txt`
2. Set `GROQ_API_KEY` from `.env.example` in your server environment.
3. Start the API: `python server.py`
4. In another terminal, serve the frontend: `python -m http.server 4173`
5. Open `http://localhost:4173`.

The browser calls `http://localhost:8000/api/ask`. The API imports the workbook into SQLite, gives Groq `execute_sql` and `custom_sql_query` tools, validates that tool calls are read-only, and returns the answer with the executed SQL evidence.
