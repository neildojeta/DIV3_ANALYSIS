import pandas as pd
import os
import logging
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import tkinter as tk
from tkinter import filedialog
import dashboard as db
import time

log_folder = "Logs"
os.makedirs(log_folder, exist_ok=True)
# Set up logging
logging.basicConfig(
    level=logging.DEBUG, 
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('Logs/Comparison.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()

# Declare global variables
file_entry_previous = None
file_entry_latest = None

prev_operator, prev_date = 0, 0
lat_operator, lat_date = 0, 0

calculated_amount = []
# calculated_totals = 0

def load_sheets(file_previous, file_latest):
    try:
        logger.info("Loading sheets from the provided Excel files.")

        # Define possible sheet names
        sheet_options = {
            "PR": ["PR"],
            "ADA_Hours": ["ADAHours"],
            "ADA": ["ADA"],
            "GOLINK_Hours": ["GOLINKHours"],
            "GOLINK": ["GOLINK"],
            "StandbyADAGOLINK": ["StandbyADAGOLINK"],
            "Operator": ["Div3PartnerList"],
            "Deductions": ["Deductions&OtherEarnings"]
        }
        
        def load_sheet(file, sheet_names):
            """Tries to load the first available sheet from a list of sheet names."""
            available_sheets = pd.ExcelFile(file).sheet_names
            for sheet in sheet_names:
                if sheet in available_sheets:
                    return pd.read_excel(file, sheet_name=sheet)
            logger.warning(f"None of the specified sheets {sheet_names} found in {file}")
            return None  # Return None if no matching sheet is found
        
        sheet_pr_previous = load_sheet(file_previous, sheet_options["PR"])
        sheet_pr_latest = load_sheet(file_latest, sheet_options["PR"])

        sheet_ADA_Hours_previous = load_sheet(file_previous, sheet_options["ADA_Hours"])
        sheet_ADA_Hours_latest = load_sheet(file_latest, sheet_options["ADA_Hours"])

        sheet_ADA_previous = load_sheet(file_previous, sheet_options["ADA"])
        sheet_ADA_latest = load_sheet(file_latest, sheet_options["ADA"])

        sheet_GOLINK_Hours_previous = load_sheet(file_previous, sheet_options["GOLINK_Hours"])
        sheet_GOLINK_Hours_latest = load_sheet(file_latest, sheet_options["GOLINK_Hours"])

        sheet_GOLINK_previous = load_sheet(file_previous, sheet_options["GOLINK"])
        sheet_GOLINK_latest = load_sheet(file_latest, sheet_options["GOLINK"])

        sheet_StandbyADAGOLINK_previous = load_sheet(file_previous, sheet_options["StandbyADAGOLINK"])
        sheet_StandbyADAGOLINK_latest = load_sheet(file_latest, sheet_options["StandbyADAGOLINK"])

        sheet_operator_previous = load_sheet(file_previous, sheet_options["Operator"])
        sheet_operator_latest = load_sheet(file_latest, sheet_options["Operator"])

        sheet_deductions_previous = load_sheet(file_previous, sheet_options["Deductions"])
        sheet_deductions_latest = load_sheet(file_latest, sheet_options["Deductions"])
        
        logger.info("Sheets loaded successfully.")
        return sheet_pr_previous, sheet_pr_latest, sheet_ADA_previous, sheet_ADA_latest, sheet_ADA_Hours_previous, sheet_ADA_Hours_latest, sheet_GOLINK_Hours_previous, sheet_GOLINK_Hours_latest, sheet_GOLINK_previous, sheet_GOLINK_latest, sheet_StandbyADAGOLINK_previous, sheet_StandbyADAGOLINK_latest, sheet_operator_previous, sheet_operator_latest, sheet_deductions_previous, sheet_deductions_latest
    except FileNotFoundError as e:
        logger.error(f"File not found: {e}")
        raise
    except Exception as e:
        logger.error(f"Error loading sheets: {e}")
        raise

def clean_currency(value):
    try:
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').strip()
            return round(float(value), 2) if value else None
        return value
    except ValueError:
        logger.error(f"Error cleaning currency value: {value}")
        return None

def calculate_totals(deductions_sheet, pr_sheet):
    try:
        if not isinstance(deductions_sheet, pd.DataFrame):
            raise TypeError("deductions_sheet should be a pandas DataFrame.")
        if not isinstance(pr_sheet, pd.DataFrame):
            raise TypeError("pr_sheet should be a pandas DataFrame.")

        calculated_totals = 0

        client_header_index = 3  # Manually setting starting point

        # Find the first empty row after the client header row
        empty_row_index = pr_sheet.iloc[client_header_index + 1:, 0].isna().idxmax()

        if pd.isna(empty_row_index):
            next_header_index = pr_sheet.shape[0]
        else:
            next_header_index = empty_row_index + client_header_index + 1

        # Extract the partner rows
        partner_rows = pr_sheet.iloc[client_header_index + 1:next_header_index]

        # Filter out rows where the Partner (first column) is empty or 0
        valid_partner_rows = partner_rows[
            partner_rows.iloc[:, 0].notna() & (partner_rows.iloc[:, 0] != 0)
        ]

        # Get partner names as a list
        partner_names = valid_partner_rows.iloc[:, 0].astype(str).str.strip().tolist()

        logger.info(f"Total valid partners: {len(valid_partner_rows)}")
        logger.info(f"Partner names: {partner_names}")

        total_amount = 0

        # Sum the Amount column (column 13, index 13) for valid partners
        total_amount = valid_partner_rows.iloc[:, 13].sum()

        calculated_totals += total_amount

        logger.info(f"Total amount: {total_amount}")
        logger.info(f"Total calculated amount: {calculated_totals}")

        return calculated_totals

    except Exception as e:
        logger.error(f"Error calculating totals: {e}")
        raise

def compare_totals(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing Totals between previous and latest values.")

        # Ensure inputs are numeric
        if not isinstance(sheet_previous, (int, float)) or not isinstance(sheet_latest, (int, float)):
            raise TypeError("Both sheet_previous and sheet_latest must be numeric values.")
        
        # Create a DataFrame to store results
        deductions_comparison = pd.DataFrame({
            "LATEST": [sheet_latest],
            "PREVIOUS": [sheet_previous],
            "DIFFERENCE": [sheet_latest - sheet_previous]
        }).round(2)

        # Add CHANGE column based on the difference
        deductions_comparison["CHANGE"] = deductions_comparison["DIFFERENCE"].apply(
            lambda diff: "Increased" if diff > 0 else "Decreased" if diff < 0 else "No Change"
        )

        logger.info("Totals comparison completed.")
        return deductions_comparison

    except Exception as e:
        logger.error(f"Error comparing totals: {e}")
        raise

def compare_TTL_Rev(sheet_previous, sheet_latest):
    try:
        if not isinstance(sheet_previous, pd.DataFrame):
            raise TypeError("sheet_previous should be a pandas DataFrame.")
        if not isinstance(sheet_latest, pd.DataFrame):
            raise TypeError("sheet_latest should be a pandas DataFrame.")
        
        sheet_previous = sheet_previous.dropna(subset=[sheet_previous.columns[0], sheet_previous.columns[3]])
        sheet_latest = sheet_latest.dropna(subset=[sheet_latest.columns[0], sheet_latest.columns[3]])

        previous_data = sheet_previous.iloc[:, [0, 3]].copy()
        latest_data = sheet_latest.iloc[:, [0, 3]].copy()

        previous_data.columns = ["PARTNER", "PREVIOUS"]
        latest_data.columns = ["PARTNER", "LATEST"]

        comparison = pd.merge(latest_data, previous_data, on="PARTNER", how="outer")

        comparison["LATEST"] = comparison["LATEST"].fillna(0)
        comparison["PREVIOUS"] = comparison["PREVIOUS"].fillna(0)

        # Convert columns to numeric
        comparison["LATEST"] = pd.to_numeric(comparison["LATEST"], errors='coerce').fillna(0)
        comparison["PREVIOUS"] = pd.to_numeric(comparison["PREVIOUS"], errors='coerce').fillna(0)

        comparison["CHANGE"] = comparison["LATEST"] - comparison["PREVIOUS"]

        comparison = comparison[["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]]

        comparison = comparison.drop(index=0)

        return comparison

    except Exception as e:
        logger.error(f"Error comparing TTL Rev: {e}")
        raise

def compare_liftlease(sheet_previous, sheet_latest):
    try:
        if not isinstance(sheet_previous, pd.DataFrame):
            raise TypeError("sheet_previous should be a pandas DataFrame.")
        if not isinstance(sheet_latest, pd.DataFrame):
            raise TypeError("sheet_latest should be a pandas DataFrame.")
        
        sheet_previous = sheet_previous.dropna(subset=[sheet_previous.columns[0], sheet_previous.columns[9]])
        sheet_latest = sheet_latest.dropna(subset=[sheet_latest.columns[0], sheet_latest.columns[9]])

        previous_data = sheet_previous.iloc[:, [0, 9]].copy()
        latest_data = sheet_latest.iloc[:, [0, 9]].copy()

        previous_data.columns = ["PARTNER", "PREVIOUS"]
        latest_data.columns = ["PARTNER", "LATEST"]

        comparison = pd.merge(latest_data, previous_data, on="PARTNER", how="outer")

        comparison["LATEST"] = comparison["LATEST"].fillna(0)
        comparison["PREVIOUS"] = comparison["PREVIOUS"].fillna(0)

        # Convert columns to numeric
        comparison["LATEST"] = pd.to_numeric(comparison["LATEST"], errors='coerce').fillna(0)
        comparison["PREVIOUS"] = pd.to_numeric(comparison["PREVIOUS"], errors='coerce').fillna(0)

        comparison["CHANGE"] = comparison["LATEST"] - comparison["PREVIOUS"]

        comparison = comparison[["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]]

        comparison = comparison.drop(index=0)

        return comparison

    except Exception as e:
        logger.error(f"Error comparing Lift Lease: {e}")
        raise

def compare_violations(sheet_previous, sheet_latest):
    try:
        if not isinstance(sheet_previous, pd.DataFrame):
            raise TypeError("sheet_previous should be a pandas DataFrame.")
        if not isinstance(sheet_latest, pd.DataFrame):
            raise TypeError("sheet_latest should be a pandas DataFrame.")
        
        sheet_previous = sheet_previous.dropna(subset=[sheet_previous.columns[0], sheet_previous.columns[10]])
        sheet_latest = sheet_latest.dropna(subset=[sheet_latest.columns[0], sheet_latest.columns[10]])

        previous_data = sheet_previous.iloc[:, [0, 10]].copy()
        latest_data = sheet_latest.iloc[:, [0, 10]].copy()

        previous_data.columns = ["PARTNER", "PREVIOUS"]
        latest_data.columns = ["PARTNER", "LATEST"]

        comparison = pd.merge(latest_data, previous_data, on="PARTNER", how="outer")

        comparison["LATEST"] = comparison["LATEST"].fillna(0)
        comparison["PREVIOUS"] = comparison["PREVIOUS"].fillna(0)

        # Convert columns to numeric
        comparison["LATEST"] = pd.to_numeric(comparison["LATEST"], errors='coerce').fillna(0)
        comparison["PREVIOUS"] = pd.to_numeric(comparison["PREVIOUS"], errors='coerce').fillna(0)

        comparison["CHANGE"] = comparison["LATEST"] - comparison["PREVIOUS"]

        comparison = comparison[["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]]

        comparison = comparison.drop(index=0)

        return comparison

    except Exception as e:
        logger.error(f"Error comparing Violations: {e}")
        raise

def compare_cash_collected(sheet_previous, sheet_latest):
    try:
        if not isinstance(sheet_previous, pd.DataFrame):
            raise TypeError("sheet_previous should be a pandas DataFrame.")
        if not isinstance(sheet_latest, pd.DataFrame):
            raise TypeError("sheet_latest should be a pandas DataFrame.")
        
        sheet_previous = sheet_previous.dropna(subset=[sheet_previous.columns[0], sheet_previous.columns[11]])
        sheet_latest = sheet_latest.dropna(subset=[sheet_latest.columns[0], sheet_latest.columns[11]])

        previous_data = sheet_previous.iloc[:, [0, 11]].copy()
        latest_data = sheet_latest.iloc[:, [0, 11]].copy()

        previous_data.columns = ["PARTNER", "PREVIOUS"]
        latest_data.columns = ["PARTNER", "LATEST"]

        comparison = pd.merge(latest_data, previous_data, on="PARTNER", how="outer")

        comparison["LATEST"] = comparison["LATEST"].fillna(0)
        comparison["PREVIOUS"] = comparison["PREVIOUS"].fillna(0)

        # Convert columns to numeric
        comparison["LATEST"] = pd.to_numeric(comparison["LATEST"], errors='coerce').fillna(0)
        comparison["PREVIOUS"] = pd.to_numeric(comparison["PREVIOUS"], errors='coerce').fillna(0)

        comparison["CHANGE"] = comparison["LATEST"] - comparison["PREVIOUS"]

        comparison = comparison[["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]]

        comparison = comparison.drop(index=0)

        return comparison

    except Exception as e:
        logger.error(f"Error comparing Cash Collected: {e}")
        raise

def compare_operators(sheet_previous, sheet_latest):
    try:
        global prev_operator, lat_operator
        logger.info("Comparing operators between previous and latest sheets.")
        
        # Extract and store unique pairs of (OPERATOR NAME, PARTNER NAME)
        operators_previous = set(sheet_previous[["Operator Name per Trip Revenue", "PARTNER"]].dropna().itertuples(index=False, name=None))
        operators_latest = set(sheet_latest[["Operator Name per Trip Revenue", "PARTNER"]].dropna().itertuples(index=False, name=None))

        # Identify added and removed operators
        added = operators_latest - operators_previous
        removed = operators_previous - operators_latest

        # Convert to DataFrame format
        added_list = [{"OPERATOR NAME": op, "PARTNER": partner, "CHANGE": "Added"} for op, partner in added]
        removed_list = [{"OPERATOR NAME": op, "PARTNER": partner, "CHANGE": "Removed"} for op, partner in removed]

        # Create DataFrames
        added_df = pd.DataFrame(added_list)
        removed_df = pd.DataFrame(removed_list)

        # Ensure DataFrames always have the necessary columns
        if added_df.empty:
            added_df = pd.DataFrame(columns=["OPERATOR NAME", "PARTNER", "CHANGE"])
        if removed_df.empty:
            removed_df = pd.DataFrame(columns=["OPERATOR NAME", "PARTNER", "CHANGE"])

        # Combine the results
        operator_changes_df = pd.concat([added_df, removed_df], ignore_index=True)

        if operator_changes_df.empty:
            operator_changes_df = pd.DataFrame({"OPERATOR NAME": ["No changes"], "PARTNER": ["No changes"], "CHANGE": ["No changes"]})

        prev_operator = sheet_previous["Operator Name per Trip Revenue"].nunique()
        lat_operator = sheet_latest["Operator Name per Trip Revenue"].nunique()
        logger.info("Operator comparison completed.")
        return operator_changes_df

    except Exception as e:
        logger.error(f"Error comparing operators: {e}")
        raise
def compare_dates(sheet_previous, sheet_latest):
    try:
        global prev_date, lat_date
        logger.info("Comparing dates between previous and latest sheets.")

        # Extracting the Date column and dropping NaN values
        dates_previous = set(sheet_previous["Date"].dropna())
        dates_latest = set(sheet_latest["Date"].dropna())

        # Identifying added and removed dates
        added_dates = dates_latest - dates_previous
        removed_dates = dates_previous - dates_latest

        # Formatting the results as lists of dictionaries
        added_list = [{"Date": date} for date in added_dates]
        removed_list = [{"Date": date} for date in removed_dates]

        prev_date = sheet_previous["Date"].nunique()
        lat_date = sheet_latest["Date"].nunique()
        logger.info("Date comparison completed.")
        return {"Added": added_list, "Removed": removed_list}
    except Exception as e:
        logger.error(f"Error comparing dates: {e}")
        raise
    
def find_missing_dates(sheet_previous, sheet_latest):
    try:
        logger.info("Checking for missing dates in the consecutive range.")
        
        # Extracting the Date column and dropping NaN values
        dates_previous = set(pd.to_datetime(sheet_previous["Date"].dropna()))
        dates_latest = set(pd.to_datetime(sheet_latest["Date"].dropna()))
        
        # Finding the full expected date range
        all_dates = pd.date_range(min(dates_previous.union(dates_latest)), 
                                  max(dates_previous.union(dates_latest)))
        
        # Finding missing dates
        missing_dates = sorted(set(all_dates) - dates_previous - dates_latest)
        
        # Converting to DataFrame
        if not missing_dates:
            missing_df = pd.DataFrame({"Missing Dates of Work": ["There are no missing dates of work."]})
        else:
            # missing_df = pd.DataFrame(missing_dates, columns=["Missing Dates"])
            missing_df = pd.DataFrame([date.strftime("-----  %B %d, %Y  -----") for date in missing_dates], columns=["Missing Dates of Work"])
        
        logger.info("Missing date check completed.")
        return missing_df
    except Exception as e:
        logger.error(f"Error finding missing dates: {e}")
        raise

def compare_Dec_Hours(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing Payable Online Hours for NIGHT AND WEEKEND operators.")

        # Filter for "NIGHT AND WEEKEND" operators
        sheet_previous_filtered = sheet_previous[sheet_previous["TYPE OF OPERATOR"] == "NIGHT AND WEEKEND"]
        sheet_latest_filtered = sheet_latest[sheet_latest["TYPE OF OPERATOR"] == "NIGHT AND WEEKEND"]

        # Group by PARTNER NAME and sum each metric
        prev_grouped = sheet_previous_filtered.groupby("PARTNER", as_index=False)["Decimal Pay Hours"].sum()
        latest_grouped = sheet_latest_filtered.groupby("PARTNER", as_index=False)["Decimal Pay Hours"].sum()

        # Merge the grouped dataframes
        comparison = latest_grouped.merge(
            prev_grouped, on="PARTNER", how="outer", suffixes=("_LATEST", "_PREVIOUS")
        ).fillna(0)

       # Calculate the change
        comparison["CHANGE"] = comparison["Decimal Pay Hours_LATEST"] - comparison["Decimal Pay Hours_PREVIOUS"]

        # Format columns as percentages with 4 significant digits
        for col in ["Decimal Pay Hours_LATEST", "Decimal Pay Hours_PREVIOUS", "CHANGE"]:
            comparison[col] = comparison[col].round(2)

        # Rename columns
        comparison.columns = ["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]

        logger.info("Comparison of Payable Online Hours and other metrics completed.")
        return comparison

    except Exception as e:
        logger.error(f"Error comparing Payable Online Hours: {e}")
        raise

def compare_Fares_Collected(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing Fares Collected for NIGHT AND WEEKEND operators.")

        # Filter for "NIGHT AND WEEKEND" operators
        sheet_previous_filtered = sheet_previous[sheet_previous["TYPE OF OPERATOR"] == "NIGHT AND WEEKEND"]
        sheet_latest_filtered = sheet_latest[sheet_latest["TYPE OF OPERATOR"] == "NIGHT AND WEEKEND"]

        # Group by PARTNER NAME and sum each metric
        prev_grouped = sheet_previous_filtered.groupby("PARTNER", as_index=False)["Fares Collected"].sum()
        latest_grouped = sheet_latest_filtered.groupby("PARTNER", as_index=False)["Fares Collected"].sum()

        # Merge the grouped dataframes
        comparison = latest_grouped.merge(
            prev_grouped, on="PARTNER", how="outer", suffixes=("_LATEST", "_PREVIOUS")
        ).fillna(0)

       # Calculate the change
        comparison["CHANGE"] = comparison["Fares Collected_LATEST"] - comparison["Fares Collected_PREVIOUS"]

        # Format columns as percentages with 4 significant digits
        for col in ["Fares Collected_LATEST", "Fares Collected_PREVIOUS", "CHANGE"]:
            comparison[col] = comparison[col].round(2)

        # Rename columns
        comparison.columns = ["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]

        logger.info("Comparison of Fares Collected completed.")
        return comparison

    except Exception as e:
        logger.error(f"Error comparing Fares Collected: {e}")
        raise

def compare_Tickets_Collected(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing Tickets Collected for NIGHT AND WEEKEND operators.")

        # Filter for "NIGHT AND WEEKEND" operators
        sheet_previous_filtered = sheet_previous[sheet_previous["TYPE OF OPERATOR"] == "NIGHT AND WEEKEND"]
        sheet_latest_filtered = sheet_latest[sheet_latest["TYPE OF OPERATOR"] == "NIGHT AND WEEKEND"]

        # Group by PARTNER NAME and sum each metric
        prev_grouped = sheet_previous_filtered.groupby("PARTNER", as_index=False)["Tickets Collected"].sum()
        latest_grouped = sheet_latest_filtered.groupby("PARTNER", as_index=False)["Tickets Collected"].sum()

        # Merge the grouped dataframes
        comparison = latest_grouped.merge(
            prev_grouped, on="PARTNER", how="outer", suffixes=("_LATEST", "_PREVIOUS")
        ).fillna(0)

       # Calculate the change
        comparison["CHANGE"] = comparison["Tickets Collected_LATEST"] - comparison["Tickets Collected_PREVIOUS"]

        # Format columns as percentages with 4 significant digits
        for col in ["Tickets Collected_LATEST", "Tickets Collected_PREVIOUS", "CHANGE"]:
            comparison[col] = comparison[col].round(2)

        # Rename columns
        comparison.columns = ["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]

        logger.info("Comparison of Tickets Collected completed.")
        return comparison

    except Exception as e:
        logger.error(f"Error comparing Tickets Collected: {e}")
        raise

def compare_Booking(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing Booking Count between previous and latest sheets.")

        sheet_previous_filtered = sheet_previous[sheet_previous["TYPE OF OPERATOR"] == "DAILY"]
        sheet_latest_filtered = sheet_latest[sheet_latest["TYPE OF OPERATOR"] == "DAILY"]

        # Group by "PARTNER NAME" 
        prev_values = sheet_previous_filtered.groupby("PARTNER", as_index=False)["Booking ID"].count()
        latest_values = sheet_latest_filtered.groupby("PARTNER", as_index=False)["Booking ID"].count()

        # Merge both datasets
        comparison = latest_values.merge(
            prev_values, on="PARTNER", how="outer", suffixes=("_LATEST", "_PREVIOUS")
        ).fillna(0)

        # Calculate the change
        comparison["CHANGE"] = comparison["Booking ID_LATEST"] - comparison["Booking ID_PREVIOUS"]

        # Format columns as percentages with 4 significant digits
        for col in ["Booking ID_LATEST", "Booking ID_PREVIOUS", "CHANGE"]:
            comparison[col] = comparison[col].round(2)

        # Rename columns
        comparison.columns = ["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]

        logger.info("Booking Count comparison completed.")
        return comparison
    except Exception as e:
        logger.error(f"Error comparing Booking Count: {e}")
        raise

def compare_Per_Trip_Revenue(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing Booking Count between previous and latest sheets.")

        sheet_previous_filtered = sheet_previous[sheet_previous["TYPE OF OPERATOR"] == "DAILY"]
        sheet_latest_filtered = sheet_latest[sheet_latest["TYPE OF OPERATOR"] == "DAILY"]

        # Group by "PARTNER NAME" 
        prev_values = sheet_previous_filtered.groupby("PARTNER", as_index=False)["Per Trip Revenue"].count()
        latest_values = sheet_latest_filtered.groupby("PARTNER", as_index=False)["Per Trip Revenue"].count()

        # Merge both datasets
        comparison = latest_values.merge(
            prev_values, on="PARTNER", how="outer", suffixes=("_LATEST", "_PREVIOUS")
        ).fillna(0)

        # Calculate the change
        comparison["CHANGE"] = comparison["Per Trip Revenue_LATEST"] - comparison["Per Trip Revenue_PREVIOUS"]

        # Format columns as percentages with 4 significant digits
        for col in ["Per Trip Revenue_LATEST", "Per Trip Revenue_PREVIOUS", "CHANGE"]:
            comparison[col] = comparison[col].round(2)

        # Rename columns
        comparison.columns = ["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]

        logger.info("Per Trip Revenue comparison completed.")
        return comparison
    except Exception as e:
        logger.error(f"Error comparing Per Trip Revenue: {e}")
        raise

def compare_Per_Trip_Revenue(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing Booking Count between previous and latest sheets.")

        sheet_previous_filtered = sheet_previous[sheet_previous["TYPE OF OPERATOR"] == "DAILY"]
        sheet_latest_filtered = sheet_latest[sheet_latest["TYPE OF OPERATOR"] == "DAILY"]

        # Group by "PARTNER NAME" 
        prev_values = sheet_previous_filtered.groupby("PARTNER", as_index=False)["Per Trip Revenue"].sum()
        latest_values = sheet_latest_filtered.groupby("PARTNER", as_index=False)["Per Trip Revenue"].sum()

        # Merge both datasets
        comparison = latest_values.merge(
            prev_values, on="PARTNER", how="outer", suffixes=("_LATEST", "_PREVIOUS")
        ).fillna(0)

        # Calculate the change
        comparison["CHANGE"] = comparison["Per Trip Revenue_LATEST"] - comparison["Per Trip Revenue_PREVIOUS"]

        # Format columns as percentages with 4 significant digits
        for col in ["Per Trip Revenue_LATEST", "Per Trip Revenue_PREVIOUS", "CHANGE"]:
            comparison[col] = comparison[col].round(2)

        # Rename columns
        comparison.columns = ["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]

        logger.info("Per Trip Revenue comparison completed.")
        return comparison
    except Exception as e:
        logger.error(f"Error comparing Per Trip Revenue: {e}")
        raise

def compare_tFares_Collected(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing Fare Collected for NIGHT AND WEEKEND operators.")

        # Filter for "NIGHT AND WEEKEND" operators
        sheet_previous_filtered = sheet_previous[sheet_previous["TYPE OF OPERATOR"] == "DAILY"]
        sheet_latest_filtered = sheet_latest[sheet_latest["TYPE OF OPERATOR"] == "DAILY"]

        # Group by PARTNER NAME and sum each metric
        prev_grouped = sheet_previous_filtered.groupby("PARTNER", as_index=False)["Fare Collected"].sum()
        latest_grouped = sheet_latest_filtered.groupby("PARTNER", as_index=False)["Fare Collected"].sum()

        # Merge the grouped dataframes
        comparison = latest_grouped.merge(
            prev_grouped, on="PARTNER", how="outer", suffixes=("_LATEST", "_PREVIOUS")
        ).fillna(0)

       # Calculate the change
        comparison["CHANGE"] = comparison["Fare Collected_LATEST"] - comparison["Fare Collected_PREVIOUS"]

        # Format columns as percentages with 4 significant digits
        for col in ["Fare Collected_LATEST", "Fare Collected_PREVIOUS", "CHANGE"]:
            comparison[col] = comparison[col].round(2)

        # Rename columns
        comparison.columns = ["PARTNER", "LATEST", "PREVIOUS", "CHANGE"]

        logger.info("Comparison of Fare Collected completed.")
        return comparison

    except Exception as e:
        logger.error(f"Error comparing Fare Collected: {e}")
        raise

def apply_formatting(sheet_name, wb):
    try:
        logger.info(f"Applying formatting to sheet: {sheet_name}.")
        ws = wb[sheet_name]
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws.columns:
            max_length = max(len(str(cell.value)) for cell in col if cell.value)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2

        # for col in ws.columns:
        #     max_length = 0
        #     column = col[0].column_letter
        #     for cell in col:
        #         try:
        #             if cell.value:
        #                 cell_length = len(str(cell.value))
        #                 if cell_length > max_length:
        #                     max_length = cell_length
        #         except:
        #             pass
        #     if ws[column + '1'].value:
        #         header_length = len(str(ws[column + '1'].value))
        #         if header_length > max_length:
        #             max_length = header_length

        #     adjusted_width = (max_length + 0.5) # add scaling
        #     ws.column_dimensions[column].width = adjusted_width

        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border
                if ws[1][cell.column - 1].value.lower() == "change":
                    if isinstance(cell.value, (int, float)):
                        if cell.value > 0:
                            cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                            cell.font = Font(color="006100")
                        elif cell.value < 0:
                            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                            cell.font = Font(color="9C0006")
                    elif isinstance(cell.value, str):
                        if cell.value.lower() in ["increased", "added"]:
                            cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                            cell.font = Font(color="006100")
                        elif cell.value.lower() in ["decreased", "removed"]:
                            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                            cell.font = Font(color="9C0006")
        logger.info(f"Formatting applied successfully to sheet: {sheet_name}.")
    except Exception as e:
        logger.error(f"Error applying formatting to sheet {sheet_name}: {e}")
        raise


def save_comparison_results(output_folder, comparison_data, filename):
    try:
        logger.info(f"Saving comparison results to {filename}.")
        os.makedirs(output_folder, exist_ok=True)
        full_comparison_file = os.path.join(output_folder, filename)
        with pd.ExcelWriter(full_comparison_file, engine="openpyxl") as writer:
            for sheet_name, data in comparison_data.items():
                data.to_excel(writer, sheet_name=sheet_name, index=False)

        wb_full = load_workbook(full_comparison_file)
        for sheet in comparison_data.keys():
            apply_formatting(sheet, wb_full)
        wb_full.save(full_comparison_file)
        wb_full.close()

        logger.info(f"Comparison results saved successfully to {filename}.")
    except Exception as e:
        logger.error(f"Error saving comparison results: {e}")
        raise

def main(file_previous, file_latest):
    try: 
        file_entry_previous = file_previous
        file_entry_latest = file_latest

        print(f"{file_entry_previous} + {file_entry_latest}")
        logger.info("Starting main comparison process.")
        output_folder = "ComparedResults"
        os.makedirs(output_folder, exist_ok=True)

        sheet_pr_previous, sheet_pr_latest, sheet_ADA_previous, sheet_ADA_latest, sheet_ADA_Hours_previous, sheet_ADA_Hours_latest, sheet_GOLINK_Hours_previous, sheet_GOLINK_Hours_latest, sheet_GOLINK_previous, sheet_GOLINK_latest, sheet_StandbyADAGOLINK_previous, sheet_StandbyADAGOLINK_latest, sheet_operator_previous, sheet_operator_latest, sheet_deductions_previous, sheet_deductions_latest = load_sheets(file_entry_previous, file_entry_latest)
# 1. Process the data without any filtering (full comparison)

        # 1. Compare totals
        prev_totals = calculate_totals(sheet_deductions_previous, sheet_pr_previous)
        lat_totals = calculate_totals(sheet_deductions_latest, sheet_pr_latest)
        totals_comparison_df = compare_totals(prev_totals, lat_totals)
        logger.info(f"\nTotals DF: {totals_comparison_df}")

        # Save the Hourly TTL Rev comparison results
        ttl_rev_comparison_df = compare_TTL_Rev(sheet_pr_previous, sheet_pr_latest)
        logger.info(f"Hourly TTL Rev Comparison DF: {ttl_rev_comparison_df}")

        # # Save the Per Trip Revenue comparison results
        # per_trip_revenue_comparison_df = compare_Per_Trip_Revenue(sheet_trips_previous, sheet_trips_latest)
        # logger.info(f"Per Trip Revenue Comparison DF: {per_trip_revenue_comparison_df}")

        # # 2. Compare Lift Lease
        compare_liftlease_df = compare_liftlease(sheet_pr_previous, sheet_pr_latest)
        logger.info(f"Lift Lease DF: {compare_liftlease_df}")

        # # 3. Compare Violations
        compare_violations_df = compare_violations(sheet_pr_previous, sheet_pr_latest)
        logger.info(f"Violations DF: {compare_violations_df}")

        # 4. Compare Cash Collected
        compare_cash_collected_df = compare_cash_collected(sheet_pr_previous, sheet_pr_latest)
        logger.info(f"Cash Collected DF: {compare_cash_collected_df}")

        # # 4. Compare operators
        # operator_changes_df = compare_operators(sheet_operator_previous, sheet_operator_latest)
        # logger.info(f"Operator Changes DF:{operator_changes_df}")

        # # Compare dates
        # date_changes = compare_dates(sheet_hours_previous, sheet_hours_latest)
        # Dadded_df = pd.DataFrame(date_changes["Added"])
        # Dremoved_df = pd.DataFrame(date_changes["Removed"])
        # Dadded_df["Change"] = "Added"
        # Dremoved_df["Change"] = "Removed"
        # Doperator_changes_df = pd.concat([Dadded_df, Dremoved_df], ignore_index=True)
        # # 5. Compare Dates
        # missing_dates_df = find_missing_dates(sheet_hours_previous, sheet_hours_latest)
        # logger.info(f"Missing Dates DF:{missing_dates_df}")

        # Save the full comparison results

        full_comparison_file = os.path.join(output_folder, "DIV3_Main_Tables.xlsm")
        excel_sheets = []
        with pd.ExcelWriter(full_comparison_file,engine="xlsxwriter") as writer:
            # full_comparison_file = os.path.join(output_folder, "DIV3_Main_Tables.xlsx")     
            # with pd.ExcelWriter(full_comparison_file, engine="openpyxl") as writer:
            totals_comparison_df.to_excel(writer, sheet_name="TotalInvoicePayment", index=False)
            excel_sheets.append("TotalInvoicePayment")
            ttl_rev_comparison_df.to_excel(writer, sheet_name="HourlyTTLRevComparison", index=False)
            excel_sheets.append("HourlyTTLRevComparison")
            compare_liftlease_df.to_excel(writer, sheet_name="LiftLeaseComparison", index=False)
            excel_sheets.append("LiftLeaseComparison")
            compare_violations_df.to_excel(writer, sheet_name="ViolationsComparison", index=False)
            excel_sheets.append("ViolationsComparison")
            compare_cash_collected_df.to_excel(writer, sheet_name="CashCollectedComparison", index=False)
            excel_sheets.append("CashCollectedComparison")
            writer.book.add_vba_project('path_to_your_vbaProject.bin')
            writer.save()

        # # Doperator_changes_df.to_excel(writer, sheet_name="DateComparison", index=False)

        # Apply formatting to the full comparison file
        wb_full = load_workbook(full_comparison_file)
        for sheet in excel_sheets:                                                                                                
            apply_formatting(sheet, wb_full)
        wb_full.save(full_comparison_file)
        wb_full.close()

        logger.info(f"Main comparison process completed successfully. File saved to {full_comparison_file}.")

        # # Save the DecHours comparison results --------------------------
        # hours_comparison_df = compare_Dec_Hours(sheet_hours_previous, sheet_hours_latest)
        # logger.info(f"Dec Hours Comparison DF: {hours_comparison_df}")

        # # Save the Fares Collected comparison results
        # fares_collected_comparison_df = compare_Fares_Collected(sheet_hours_previous, sheet_hours_latest)
        # logger.info(f"Fares Collected Comparison DF: {fares_collected_comparison_df}")

        # # Save the Tickets Collected comparison results
        # tickets_collected_comparison_df = compare_Tickets_Collected(sheet_hours_previous, sheet_hours_latest)
        # logger.info(f"Tickets Collected Comparison DF: {tickets_collected_comparison_df}")

        # # Save the comparison results to an Excel file
        # hours_comparison_file = os.path.join(output_folder, "DIV8_Hours_Comparison.xlsx")
        # excel_sheets = []
        # with pd.ExcelWriter(hours_comparison_file, engine="openpyxl") as writer:
        #     hours_comparison_df.to_excel(writer, sheet_name="DecHoursComparison", index=False)
        #     excel_sheets.append("DecHoursComparison")
        #     fares_collected_comparison_df.to_excel(writer, sheet_name="FaresCollectedComparison", index=False)
        #     excel_sheets.append("FaresCollectedComparison")
        #     tickets_collected_comparison_df.to_excel(writer, sheet_name="TicketsCollectedComparison", index=False)
        #     excel_sheets.append("TicketsCollectedComparison")

        # # Apply formatting to the full comparison file
        # wb_full = load_workbook(hours_comparison_file)
        # for sheet in excel_sheets:                                                                                                
        #     apply_formatting(sheet, wb_full)
        # wb_full.save(hours_comparison_file)
        # wb_full.close()

        # logger.info(f"Hours comparison process completed successfully. File saved to {hours_comparison_file}.")

        # # Save the Booking Count comparison results
        # booking_count_comparison_df = compare_Booking(sheet_trips_previous, sheet_trips_latest)
        # logger.info(f"Booking Count Comparison DF: {booking_count_comparison_df}")
        
        # # Save Fares Collected comparison results
        # tfares_collected_comparison_df = compare_tFares_Collected(sheet_trips_previous, sheet_trips_latest)
        # logger.info(f"Fares Collected Comparison DF: {tfares_collected_comparison_df}")
        
        # # Save the comparison results to an Excel file
        # trips_comparison_file = os.path.join(output_folder, "DIV8_Trips_Comparison.xlsx")
        # excel_sheets = []
        # with pd.ExcelWriter(trips_comparison_file, engine="openpyxl") as writer:
        #     booking_count_comparison_df.to_excel(writer, sheet_name="BookingCountComparison", index=False)
        #     excel_sheets.append("BookingCountComparison")
            
        #     tfares_collected_comparison_df.to_excel(writer, sheet_name="FaresCollectedComparison", index=False)
        #     excel_sheets.append("FaresCollectedComparison")
        
        # wb_full = load_workbook(trips_comparison_file)
        # for sheet in excel_sheets:
        #     apply_formatting(sheet, wb_full)
        # wb_full.save(trips_comparison_file)
        # wb_full.close()

        # logger.info(f"Trips comparison process completed successfully. File saved to {trips_comparison_file}.")

        # time.sleep(2)
        # db.main(file_previous, file_latest, prev_operator, prev_date, lat_operator, lat_date)
    except Exception as e:
        logger.error(f"Error in main comparison process: {e}")
        raise


    # finally:
    #     wb_client.close()


def open_file_dialog(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsm;*.xlsx")])
    if filename:
        entry.delete(0, tk.END)
        entry.insert(0, filename)


def create_gui():
    global file_entry_previous, file_entry_latest # Declare them as global variables

    # Set up the GUI window
    root = tk.Tk()
    root.title("Comparison Tool")

    # Create and place labels, entry boxes, and buttons
    # tk.Label(root, text="Previous File:").grid(row=0, column=0, padx=10, pady=5)
    # entry_previous = tk.Entry(root, width=50)
    # entry_previous.grid(row=0, column=1, padx=10, pady=5)
    
    # # tk.Button(root, text="Browse", command=lambda: open_file_dialog(entry_previous)).grid(row=0, column=2, padx=10, pady=5)

    # tk.Label(root, text="Latest File:").grid(row=1, column=0, padx=10, pady=5)
    # entry_latest = tk.Entry(root, width=50)
    # entry_latest.grid(row=1, column=1, padx=10, pady=5)

    prev = "VDP_DIV3_0310_0323 FINAL.xlsm"
    latest = "VDP_DIV3_0324_0406 FINAL.xlsm"
    
    # tk.Button(root, text="Browse", command=lambda: open_file_dialog(entry_latest)).grid(row=1, column=2, padx=10, pady=5)

    # Button to trigger the comparison process
    # tk.Button(root, text="Compare", command=lambda: (main(entry_previous.get(), entry_latest.get()), root.destroy())).grid(row=2, column=1, pady=20)
    # tk.Button(root, text="Compare", command=lambda: handle_comparison(entry_previous.get(), entry_latest.get(), root)).grid(row=2, column=1, pady=20)
    tk.Button(root, text="Compare", command=lambda: handle_comparison(prev, latest, root)).grid(row=2, column=1, pady=20)

    def handle_comparison(file_previous, file_latest, root):
        try:
            main(file_previous, file_latest)
        except Exception as e:
            print(f"An error occurred: {e}")
            # Check for the disconnection error and close the GUI if it happens
            if isinstance(e, OSError) and "The object invoked has disconnected" in str(e):
                print("Disconnected from Excel, closing GUI.")
                root.quit()  # This will close the Tkinter window
        finally:
            root.destroy()  # Close the window in all cases
    # Start the GUI loop
    root.mainloop()


if __name__ == "__main__":
    create_gui()

