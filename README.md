# CSV Plotter App

A Streamlit application for visualizing and analyzing temperature data from CSV or Excel files.

## Features

-   Upload CSV or Excel files.
-   Interactively trim data by time range.
-   Apply various smoothing filters (Moving Average, Savitzky-Golay).
-   Perform curve fitting and project future temperature trends.
-   Analyze temperature stability.
-   Export processed data and plots to Excel.

## How to Run

1.  **Install dependencies:**
    ```bash
    pip install streamlit pandas numpy matplotlib scipy scikit-learn openpyxl
    ```

2.  **Run the app:**
    ```bash
    streamlit run app.py
    ```
