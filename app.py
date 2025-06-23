import streamlit as st
from query_processor import FinancialQueryProcessor
from langchain_huggingface import HuggingFaceEndpoint
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
import os
import json
import logging
import tempfile
from metric_dictionary import metric_mappings, suggested_queries

# Configure logging
logging.basicConfig(filename='app.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Custom CSS for UI styling
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
        padding: 20px;
    }
    .stApp {
        max-width: 1200px;
        margin: auto;
    }
    h1 {
        color: #1e3a8a;
        font-family: 'Arial', sans-serif;
    }
    h2 {
        color: #3b82f6;
        font-family: 'Arial', sans-serif;
    }
    .stTextInput > div > div > input {
        border-radius: 8px;
        padding: 10px;
    }
    .stSelectbox > div > div > select {
        border-radius: 8px;
        padding: 10px;
    }
    .error {
        color: #dc2626;
        font-weight: bold;
    }
    .stFileUploader > div > div {
        border: 2px dashed #3b82f6;
        border-radius: 8px;
        padding: 20px;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)


# Define helper function to get unique metrics for a sheet
def get_sheet_metrics(processor, sheet_name):
    if not processor or sheet_name == "Upload Excel file first":
        return list(metric_mappings.values())
    
    # Access the sheet
    sheet = processor.wb[sheet_name]
    
    # Find header row (assuming similar logic to FinancialQueryProcessor._find_header_row)
    header_row = None
    for row in sheet.iter_rows(min_row=1, max_row=10):  # Check first 10 rows
        if any(cell.value and isinstance(cell.value, str) and cell.value.strip() in metric_mappings for cell in row):
            header_row = row
            break
    
    if not header_row:
        return []
    
    # Collect unique metrics, prioritizing leftmost occurrence
    metrics = {}
    for col_idx, cell in enumerate(header_row, start=1):
        if cell.value and isinstance(cell.value, str):
            metric = cell.value.strip()
            if metric in metric_mappings:
                # Store metric only if not seen before (leftmost occurrence)
                if metric not in metrics:
                    metrics[metric] = metric_mappings[metric]
    
    return list(metrics.values())

# Define query types for dropdown
def get_query_types(processor=None):
    sheet_names = processor.wb.sheetnames if processor else ["Upload Excel file first"]
    return [
        {
            "template": "a for b",
            "display": "Metric for Company",
            "parameters": [
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"},
                {"name": "company", "type": "text_input", "label": "Enter Company Code (e.g., ALLI)"}
            ],
            "query_format": "{metric} for {company}",
            "example": "P/E for ALLI"
        },
        {
            "template": "a and b for c",
            "display": "Two Metrics for Company",
            "parameters": [
                {"name": "metric1", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select First Metric"},
                {"name": "metric2", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Second Metric"},
                {"name": "company", "type": "text_input", "label": "Enter Company Code (e.g., ALLI)"}
            ],
            "query_format": "{metric1} and {metric2} for {company}",
            "example": "P/E and Revenue 3M for BOC"
        },
        {
            "template": "a from X and b from Y for c",
            "display": "Metrics from Specific Sheets for Company",
            "parameters": [
                {"name": "metric1", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select First Metric"},
                {"name": "sheet1", "type": "text_input", "label": "Enter First Sheet Name (e.g., Dec 2024)"},
                {"name": "metric2", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Second Metric"},
                {"name": "sheet2", "type": "text_input", "label": "Enter Second Sheet Name (e.g., Sheet1)"},
                {"name": "company", "type": "text_input", "label": "Enter Company Code (e.g., ALLI)"}
            ],
            "query_format": "{metric1} from {sheet1} and {metric2} from {sheet2} for {company}",
            "example": "C.Price from Sheet1 and P/E from Dec 2024 for ALLI"
        },
        {
            "template": "a for sector b",
            "display": "Metric for Sector",
            "parameters": [
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"},
                {"name": "sector", "type": "text_input", "label": "Enter Sector Name (e.g., FOOD, BEVERAGE & TOBACCO)"}
            ],
            "query_format": "{metric} for sector {sector}",
            "example": "Revenue 3M for sector BANKS"
        },
        {
            "template": "a for all sectors",
            "display": "Metric for All Sectors",
            "parameters": [
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "{metric} for all sectors",
            "example": "Div Yield for all sectors"
        },
        {
            "template": "average a",
            "display": "Average Metric",
            "parameters": [
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "average {metric}",
            "example": "average P/E"
        },
        {
            "template": "define a",
            "display": "Define Metric",
            "parameters": [
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "define {metric}",
            "example": "define Dividend Yield"
        },
        {
            "template": "best stock by a",
            "display": "Best Stock by Metric",
            "parameters": [
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "best stock by {metric}",
            "example": "best stock by Div Yield"
        },
        {
            "template": "best stock",
            "display": "Best Stock (Composite)",
            "parameters": [],
            "query_format": "best stock",
            "example": "best stock"
        },
        {
            "template": "best sector by a",
            "display": "Best Sector by Metric",
            "parameters": [
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "best sector by {metric}",
            "example": "best sector by Revenue 3M"
        },
        {
            "template": "best a",
            "display": "Best Metric Value",
            "parameters": [
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "best {metric}",
            "example": "best Div Yield"
        },
        {
            "template": "lowest a",
            "display": "Lowest Metric Value",
            "parameters": [
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "lowest {metric}",
            "example": "lowest Div Yield"
        },
        {
            "template": "highest a",
            "display": "Highest Metric Value",
            "parameters": [
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "highest {metric}",
            "example": "highest Div Yield"
        },
        {
            "template": "compare stocks X, Y, Z by a",
            "display": "Compare Stocks by Metric",
            "parameters": [
                {"name": "stocks", "type": "text_input", "label": "Enter Stock Codes (comma-separated, e.g., ALLI, BOC, COMB)"},
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "compare stocks {stocks} by {metric}",
            "example": "compare stocks ALLI, BOC, COMB by P/E"
        },
        {
            "template": "compare stocks X, Y, Z",
            "display": "Compare Stocks (Default Metric)",
            "parameters": [
                {"name": "stocks", "type": "text_input", "label": "Enter Stock Codes (comma-separated, e.g., ALLI, BOC, COMB)"}
            ],
            "query_format": "compare stocks {stocks}",
            "example": "compare stocks ALLI, BOC"
        },
        {
            "template": "X vs Y vs sector Z by a",
            "display": "Compare Stocks and Sector",
            "parameters": [
                {"name": "stock1", "type": "text_input", "label": "Enter First Stock Code (e.g., ALLI)", "key": "stock1"},
                {"name": "stock2", "type": "text_input", "label": "Enter Second Stock Code (e.g., BOC)", "key": "stock2"},
                {"name": "sector", "type": "text_input", "label": "Enter Sector Name (e.g., BANKS)"},
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "{stock1} vs {stock2} vs sector {sector} by {metric}",
            "example": "ALLI vs BOC vs sector BANKS by P/E"
        },
        {
            "template": "X vs sector Z by a",
            "display": "Compare Stock and Sector",
            "parameters": [
                {"name": "stock", "type": "text_input", "label": "Enter Stock Code (e.g., ALLI)"},
                {"name": "sector", "type": "text_input", "label": "Enter Sector Name (e.g., BANKS)"},
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "{stock} vs sector {sector} by {metric}",
            "example": "ALLI vs sector BANKS by Cumulative Profit"
        },
        {
            "template": "sector X vs sector Y by a",
            "display": "Compare Sectors",
            "parameters": [
                {"name": "sector1", "type": "text_input", "label": "Enter First Sector Name (e.g., BANKS)"},
                {"name": "sector2", "type": "text_input", "label": "Enter Second Sector Name (e.g., INSURANCE)"},
                {"name": "metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Metric"}
            ],
            "query_format": "sector {sector1} vs sector {sector2} by {metric}",
            "example": "sector BANKS vs sector INSURANCE by ROE"
        },
        {
            "template": "show a where b between min and max",
            "display": "Companies by Metric within Range",
            "parameters": [
                {"name": "display_metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Display Metric"},
                {"name": "filter_metric", "type": "selectbox", "options": list(metric_mappings.values()), "label": "Select Filter Metric"},
                {"name": "min_value", "type": "number_input", "label": "Enter Minimum Value", "value": 0.0},
                {"name": "max_value", "type": "number_input", "label": "Enter Maximum Value", "value": 100.0}
            ],
            "query_format": "show {display_metric} where {filter_metric} between {min_value} and {max_value}",
            "example": "show Div Yield where P/E between 10 and 20"
        }
    ]

# Streamlit app layout
with st.container():
    st.title("Financial Data Explorer")
    st.markdown(
        "Upload an Excel file (.xlsx) to query financial data. Select a query type and enter parameters, "
        "or use a suggested/custom query. Charts will be displayed for applicable queries."
    )
    st.divider()

    # File uploader
    st.subheader("Upload Excel File")
    uploaded_file = st.file_uploader("Drag and drop or click to upload an Excel file", type=["xlsx"])
    processor = None

    if uploaded_file:
        # Save the uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_file_path = tmp_file.name

        # Initialize processor
        try:
            processor = FinancialQueryProcessor(tmp_file_path)
            st.success("Excel file uploaded successfully!")
        except Exception as e:
            st.error(f"Error loading Excel file: {e}. Please ensure the file is a valid .xlsx with the expected structure.")
            logging.error(f"Failed to load Excel file: {e}")
            os.remove(tmp_file_path)
            st.stop()

        # Remove temporary file
        os.remove(tmp_file_path)

        # Set up Hugging Face API for LLM fallback
        os.environ["HUGGINGFACEHUB_API_TOKEN"] = "your-huggingface-api-token"  # Replace with your actual token
        try:
            llm = HuggingFaceEndpoint(
                repo_id="mistralai/Mixtral-8x7B-Instruct-v0.1",
                temperature=0.7,
                max_new_tokens=512
            )
        except Exception as e:
            st.warning("Unable to connect to Hugging Face API. Queries will rely on local processing only.")
            logging.warning(f"Hugging Face API setup failed: {e}")
            llm = None

        # Prompt template for LangChain
        prompt_template = PromptTemplate(
            input_variables=["query"],
            template="""
You are a financial data assistant. Parse the user query and return a JSON object with the query type and relevant entities.
Query types: company, sector, general, definition, best_stock, best_sector, best_metric, compare_stocks, multi_sheet, compare_mixed.
Extract entities: metric, company, sector, stocks, criteria, metric_sheet_pairs (list of [metric, sheet]), entities (list of {name, type: 'company' or 'sector'}).
If the query is ambiguous, suggest a clarification or map to the closest type.

Examples:
- "P/E for ALLI" -> {"type": "company", "company": "ALLI", "metrics": ["P/E"]}
- "P/E for sector BANKS" -> {"type": "sector", "sector": "BANKS", "metric": "P/E"}
- "Average P/E" -> {"type": "general", "metric": "P/E"}
- "What is P/E?" -> {"type": "definition", "metric": "P/E"}
- "Best stock by Div Yield" -> {"type": "best_stock", "criteria": "Div Yield"}
- "Best sector by P/E" -> {"type": "best_sector", "criteria": "P/E"}
- "Best P/E" -> {"type": "best_metric", "metric": "P/E"}
- "Which stocks are best ALLI, DFCC, COMB" -> {"type": "compare_stocks", "stocks": ["ALLI", "DFCC", "COMB"]}
- "C.Price and P/E for ALLI" -> {"type": "multi_sheet", "company": "ALLI", "metric_sheet_pairs": [["C.Price", "Sheet1"], ["P/E", "Dec 2024"]]}
- "ALLI vs ABL vs sector BANKS by P/E" -> {"type": "compare_mixed", "entities": [{"name": "ALLI", "type": "company"}, {"name": "ABL", "type": "company"}, {"name": "BANKS", "type": "sector"}], "metric": "P/E"}

Query: {query}
Output: JSON object
"""
        )

        # LLM chain
        llm_chain = LLMChain(prompt=prompt_template, llm=llm) if llm else None

        # Query input section
        st.subheader("Enter Your Query")

        # Query type selector
        with st.container():
            query_type = st.selectbox(
                "Query Type:",
                [""] + [qt["display"] for qt in get_query_types(processor)],
                key="query_type"
            )

        # Query parameters section
        with st.container():
            st.markdown("**Query Parameters**")
            user_inputs = {}
            constructed_query = ""

            if query_type:
                selected_query = next(qt for qt in get_query_types(processor) if qt["display"] == query_type)

                if query_type == "Companies by Metric within Range":
                    col1, col2 = st.columns(2)
                    with col1:
                        user_inputs["display_metric"] = st.selectbox(
                            "Select Display Metric",
                            selected_query["parameters"][0]["options"],
                            key="display_metric"
                        )
                    with col2:
                        user_inputs["filter_metric"] = st.selectbox(
                            "Select Filter Metric",
                            selected_query["parameters"][1]["options"],
                            key="filter_metric"
                        )

                    from metric_dictionary import metric_ranges
                    filter_metric = user_inputs["filter_metric"]
                    default_min, default_max, step = metric_ranges.get(filter_metric, (0.0, 1000.0, 1.0))

                    st.subheader("Select Range for Filter Metric")
                    col3, col4 = st.columns(2)
                    with col3:
                        use_slider = st.checkbox("Use Slider for Range", value=True, key="use_slider")
                    with col4:
                        st.write("")  # Spacer

                    if use_slider:
                        min_value, max_value = st.slider(
                            f"Select range for {filter_metric}",
                            min_value=float(default_min),
                            max_value=float(default_max),
                            value=(float(default_min), min(float(default_max), float(default_min) + step * 10)),
                            step=float(step),
                            key="range_slider"
                        )
                    else:
                        col5, col6 = st.columns(2)
                        with col5:
                            min_value = st.number_input(
                                f"Min {filter_metric}", value=float(default_min), step=float(step), key="min_input"
                            )
                        with col6:
                            max_value = st.number_input(
                                f"Max {filter_metric}", value=min(float(default_max), float(default_min) + step * 10),
                                step=float(step), key="max_input"
                            )

                    user_inputs["min_value"] = str(min_value)
                    user_inputs["max_value"] = str(max_value)
                else:
                    for param in selected_query["parameters"]:
                        if param["type"] == "selectbox":
                            user_inputs[param["name"]] = st.selectbox(param["label"], param["options"], key=param["name"])
                        elif param["type"] == "text_input":
                            user_inputs[param["name"]] = st.text_input(param["label"], key=param["name"])
                        elif param["type"] == "number_input":
                            user_inputs[param["name"]] = str(
                                st.number_input(param["label"], value=param.get("value", 0.0), key=param["name"])
                            )

                # Construct query if all inputs are provided
                if all(user_inputs.get(param["name"]) for param in selected_query["parameters"]):
                    try:
                        constructed_query = selected_query["query_format"].format(**user_inputs)
                    except KeyError:
                        st.warning("Please fill in all required parameters correctly.")
                        constructed_query = ""

        # Suggested/custom query section
        with st.container():
            st.markdown("**Suggested/Custom Query**")
            suggested_query = st.selectbox("Suggested Queries:", [""] + suggested_queries, key="suggested_query")
            user_query = st.text_input(
                "Custom Query:", value=suggested_query,
                placeholder="e.g., P/E for ALLI or ALLI vs BOC vs sector BANKS by P/E",
                key="custom_query"
            )

        st.divider()

        # Process query
        final_query = user_query or constructed_query
        if final_query and st.button("Run Query"):
            result, chart_data = processor.process_query(final_query)

            # Fallback to LLM
            if "Sorry" in result and llm_chain:
                try:
                    logging.info(f"Local processing failed for query: {final_query}. Falling back to LLM.")
                    response = llm_chain.run(query=final_query)
                    query_dict = json.loads(response.strip())
                    result, chart_data = processor.process_structured_query(query_dict)
                except json.JSONDecodeError as e:
                    logging.error(f"LLM JSON parsing failed: {e}, response: {response}")
                    result = "Error: LLM returned an invalid response. Please try a different query format."
                    chart_data = None
                except Exception as e:
                    logging.error(f"LLM processing failed: {e}")
                    result = "Error: Unable to process query with LLM. Please try a different query format."
                    chart_data = None

            # Display results
            st.subheader("Query Results")
            if "Sorry" in result or "Error" in result:
                st.markdown(f'<p class="error">{result}</p>', unsafe_allow_html=True)
            else:
                st.text_area("Result:", value=result, height=200, disabled=True)

            # Display charts
            if chart_data:
                st.subheader("Visualizations")
                if isinstance(chart_data, list):
                    cols = st.columns(2)
                    for idx, chart in enumerate(chart_data):
                        if chart:
                            with cols[idx % 2]:
                                st.plotly_chart(chart, use_container_width=True)
                else:
                    st.plotly_chart(chart_data, use_container_width=True)
    else:
        st.info("Please upload an Excel file to start querying.")
