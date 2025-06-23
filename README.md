# 📊 Financial Metric Query App and LLM

An intelligent, extensible financial analytics tool built using **Python**, **Streamlit**, and **LangChain**. This app allows users to upload Excel spreadsheets and query financial data using natural or structured language — extracting metrics for companies and sectors, comparing performance, and generating visual insights.

Queries are handled locally via a powerful custom query engine, with seamless fallback to **Mistral AI’s Mixtral-8x7B-Instruct** when required.

---

## ✨ Features

- 🔍 Natural language & structured queries
- 📈 Company & sector-level metric exploration
- 📊 Visualizations using Plotly
- 📚 Financial term definitions
- ⚙️ Excel sheet auto-parsing with robust header detection
- 🧠 AI fallback for complex or ambiguous queries using **Mistral**
- 🧮 Metric comparisons, averages, filters, and rankings

---

## 🖼️ Example Queries

```text
P/E for ALLI
Revenue 3M for sector BANKS
Cumulative Profit and Revenue for COMB
Compare stocks SAMP, HNB by Cumulative Profit
ALLI vs sector BANKS by Revenue 3M
Show Div Yield where P/E between 10 and 20
Define Free Cash Flow
Best sector by ROE
```
---

## 🧩 Architecture
```
.
├── app.py                   # Streamlit UI & LLM integration
├── query_processor.py       # Core engine for structured query parsing and processing
├── debug_header.py          # Excel header detection & structure analysis
├── metric_dictionary.py     # Metric definitions, mappings, ranges, and suggestions
├── sample_data/
│   └── Financials.xlsx      # Example Excel file
└── README.md
```

---

## ⚙️ Tech Stack


| Component     | Description                     |
| ------------- | ------------------------------- |
| Python 3.x    | Core language      		  |
| Streamlit  	| Frontend UI        		  |
| OpenPyXL   	| Excel file parsing 		  |
| Plotly     	| Data visualization 		  |
|LangChain   	| LLM orchestration            	  |
| Mistral AI 	| AI fallback for complex queries |

---

## 🧠 AI Integration

The app integrates with Mistral’s Mixtral-8x7B-Instruct using LangChain:
	•	Queries first parsed locally by query_processor.py
	•	If unrecognized or malformed, passed to Mistral via HuggingFaceEndpoint
	•	Ensures robust handling of edge-case queries

---

## 📦 Key Components & Functions
```
app.py
	•	Loads and parses uploaded Excel files
	•	Handles query input, routing, and fallback logic
	•	Integrates with LangChain/Mistral for LLM-based fallback
	•	Functions:
	•	get_sheet_metrics()
	•	get_query_types()

query_processor.py
	•	Core query processing logic
	•	Validates, normalizes, and interprets queries
	•	Computes metrics, comparisons, and filters
	•	Key functions include:
	•	process_structured_query(), process_query()
	•	_handle_company_multi_metric(), _compare_stocks(), _generate_*_chart()

debug_header.py
	•	Excel header and structure inference
	•	Handles merged headers, multiple header rows, sector detection
	•	Functions:
	•	get_merged_parent(), detect_sectors(), extract_header_signature(), etc.

metric_dictionary.py
	•	Contains:
	•	suggested_queries
	•	metric_mappings
	•	metric_definitions
	•	metric_ranges
```
---

## 🚀 Getting Started

✅ Prerequisites
	•	Python 3.8+
	•	API access to Mistral via Hugging Face (if fallback is enabled)

📚 Supported Query Types
	•	Metric for Company – e.g. P/E for ALLI
	•	Two Metrics for Company – e.g. Cumulative Profit and Revenue for COMB
	•	Metric for Sector – e.g. Revenue 3M for sector BANKS
	•	Compare Stocks/Sectors – e.g. Compare stocks ALLI, BOC by P/E
	•	Filter by Metric Ranges – e.g. Show ROE where P/E between 10 and 20
	•	Best/Top Metrics – e.g. Best sector by Revenue 3M
	•	Definitions – e.g. Define Market Cap


---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
