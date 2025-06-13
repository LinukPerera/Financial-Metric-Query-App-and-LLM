# Finance-Data-Processing-LLM
Finance Data Processing LLM

Python Libraries: ollama, chromadb, sentence-transformers, langchain-community, transformers, torch, pandas

The spreadsheet will have 3 sheets in it, two of them have a header and their data
header : Name, Code, C.Price, Foreign qty, Issue qty Mn 
header :  CODE, CODE, CMP, TTM EPS, Dividend Cash, Dividend Scrip, Total Dividend, Div Yield (Cash+Scrip), Payout Ratio 

The remaining sheet has headers and some sub headers and their data
header : CODE, CMP, Revenue 3M	 ((Dec 24.,	% )),	 PROFIT 3M 	((Dec 24., % )), EPS, Cumulative Revenue ((2024/25, 2023/24, %)), Cumulative Profit ((2024/25,	2023/24, %)), Trailing EPS, Previous ((FY Profit)), P/E, NAV, PBV, Issued Qty Mn, Assets (Bn), Equity (Bn), ROE, Dividend Cash, Dividend Scrip, Div Yield

Where ever (()) double brackets were used it means that theres a sub header under the header, % can be substituted for percentage. Columns like code, Cumulative Revenue ((2024/25)), Cumulative Profit ((2024/25)), P/E, PBV and Div Yield sometimes take up the entire row as they are used to indicate a change in category, eg: banks, insurance and will only have columns code, Cumulative Revenue ((2024/25)), Cumulative Profit ((2024/25)), P/E, PBV and Div. The column 'code' has the name of the category. The end of this sheet has a glossary of titles and mentions a few policies.

Additionally there are a few blank cells and cells that say'Results Pending' or merged blank cells while a few other cells in that row may contain information.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
