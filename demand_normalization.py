import pandas as pd
import numpy as np
import os

def process_demand_sheet(input_file):
    """
    Process the DEMAND sheet in the Excel file according to the requirements:
    1. Combine cells with values <30 with the next cell to the right
    2. Round values >0 to the nearest multiple of 10
    3. Ensure total demand for each item doesn't exceed the available quantity in the Data sheet

    Args:
        input_file (str): Path to the input Excel file

    Returns:
        str: Path to the output Excel file
    """
    try:
        # Load the Excel file
        print(f"Opening Excel file: {input_file}")

        # Get all sheet names
        xls = pd.ExcelFile(input_file)
        sheet_names = xls.sheet_names
        print(f"Available sheets: {sheet_names}")

        # Find the DEMAND sheet (case-insensitive)
        demand_sheet_name = None
        data_sheet_name = None

        for sheet_name in sheet_names:
            if sheet_name.lower() == 'demand':
                demand_sheet_name = sheet_name
            elif sheet_name.lower() == 'data':
                data_sheet_name = sheet_name

        if not demand_sheet_name:
            print("Error: No sheet named 'DEMAND' (case-insensitive) found in the Excel file.")
            return None

        if not data_sheet_name:
            print("Warning: No sheet named 'DATA' (case-insensitive) found in the Excel file.")
            print("The quantity constraint check will be skipped.")

        print(f"Found sheets: Demand={demand_sheet_name}, Data={data_sheet_name}")

        # Read the DEMAND sheet into a DataFrame
        df_demand = pd.read_excel(input_file, sheet_name=demand_sheet_name)

        print(f"Demand DataFrame shape: {df_demand.shape}")
        print("Demand Column headers:")
        print(df_demand.columns.tolist())

        # Read the DATA sheet if it exists
        if data_sheet_name:
            df_data = pd.read_excel(input_file, sheet_name=data_sheet_name)
            print(f"Data DataFrame shape: {df_data.shape}")
            print("Data Column headers:")
            print(df_data.columns.tolist())

            # Create a dictionary to store the total available quantity for each material
            material_qty_dict = {}

            # Check if the required columns exist in the Data sheet
            if 'Material' in df_data.columns and 'Still to be delivered (qty)' in df_data.columns:
                # Convert Material column to string to ensure consistent comparison
                df_data['Material'] = df_data['Material'].astype(str)

                # Group by Material and sum the quantities
                material_qty = df_data.groupby('Material')['Still to be delivered (qty)'].sum()
                material_qty_dict = material_qty.to_dict()

                print("Available quantities by material:")
                for material, qty in material_qty_dict.items():
                    print(f"{material}: {qty}")
            else:
                print("Warning: Required columns not found in Data sheet. Quantity constraint check will be skipped.")
                material_qty_dict = {}
        else:
            material_qty_dict = {}

        # Convert numeric date columns to proper date format
        date_columns = df_demand.columns[2:]  # Columns from C onwards
        date_columns_dict = {}

        for col in date_columns:
            if isinstance(col, (int, float)):
                # Convert Excel date number to datetime
                date_str = pd.to_datetime(col, unit='D', origin='1899-12-30').strftime('%m/%d/%Y')
                date_columns_dict[col] = date_str

        # Rename columns if date_columns_dict is not empty
        if date_columns_dict:
            df_demand = df_demand.rename(columns=date_columns_dict)
            print("Converted date columns:")
            print(date_columns_dict)

        # Process the data starting from column C (index 2)
        # Make a copy of the DataFrame to avoid modifying the original
        processed_df = df_demand.copy()

        # Get the numeric columns (from column C onwards)
        numeric_cols = processed_df.columns[2:]

        # Tạo bản sao của DataFrame gốc để lưu giá trị ban đầu
        original_df = df_demand.copy()

        # Step 1: Kiểm tra tổng nhu cầu so với số lượng có sẵn trước khi xử lý
        if material_qty_dict:
            # Assume the first column is the item code/material
            item_col = processed_df.columns[0]

            # Tạo danh sách các mã hàng cần giữ nguyên giá trị
            items_to_keep_original = []

            # For each row in the processed DataFrame
            for row_idx in range(len(processed_df)):
                item_code = processed_df.at[row_idx, item_col]

                # Convert item_code to string to match with material_qty_dict keys
                item_code_str = str(item_code)

                # If the item exists in the material_qty_dict
                if item_code_str in material_qty_dict:
                    available_qty = material_qty_dict[item_code_str]

                    # Tính tổng nhu cầu ban đầu (trước khi xử lý)
                    original_demand = original_df.iloc[row_idx, 2:].sum()

                    print(f"Mã hàng {item_code_str}: Số lượng có sẵn={available_qty}, Nhu cầu ban đầu={original_demand}")

                    # Nếu tổng nhu cầu ban đầu vượt quá số lượng có sẵn
                    if original_demand > available_qty:
                        print(f"Cảnh báo: Tổng nhu cầu ban đầu ({original_demand}) cho mã hàng {item_code_str} vượt quá số lượng có sẵn ({available_qty})")
                        print(f"Giữ nguyên giá trị cho mã hàng {item_code_str} theo yêu cầu")
                        items_to_keep_original.append(item_code_str)

        # In ra danh sách các mã hàng cần giữ nguyên
        print(f"Danh sách các mã hàng cần giữ nguyên: {items_to_keep_original}")

        # Step 2: Combine cells with values <30 (chỉ áp dụng cho các mã hàng không nằm trong danh sách giữ nguyên)
        for row_idx in range(len(processed_df)):
            item_code = str(processed_df.at[row_idx, item_col])

            # In ra thông tin để debug
            print(f"Đang xử lý mã hàng: {item_code}, Có trong danh sách giữ nguyên: {item_code in items_to_keep_original}")

            # Nếu mã hàng không nằm trong danh sách giữ nguyên
            if item_code not in items_to_keep_original:
                # In ra thông tin để debug
                print(f"Áp dụng bước gộp cho mã hàng: {item_code}")

                for col_idx in range(len(numeric_cols) - 1):  # Skip the last column
                    col_name = numeric_cols[col_idx]
                    next_col_name = numeric_cols[col_idx + 1]

                    # Check if current cell value is less than 30 and greater than 0
                    current_value = processed_df.at[row_idx, col_name]

                    # Chuyển đổi current_value thành số nếu có thể
                    try:
                        if pd.isna(current_value):
                            current_value = 0
                        else:
                            current_value = float(current_value)
                    except (ValueError, TypeError):
                        current_value = 0

                    if 0 < current_value < 30:
                        # Add current cell value to next cell
                        next_value = processed_df.at[row_idx, next_col_name]

                        # Chuyển đổi next_value thành số nếu có thể
                        try:
                            if pd.isna(next_value):
                                next_value = 0
                            else:
                                next_value = float(next_value)
                        except (ValueError, TypeError):
                            next_value = 0

                        # Gộp giá trị
                        processed_df.at[row_idx, next_col_name] = next_value + current_value

                        # Set current cell to 0
                        processed_df.at[row_idx, col_name] = 0

                        # In ra thông tin để debug
                        print(f"Đã gộp giá trị {current_value} từ cột {col_name} vào cột {next_col_name}")

        # Step 3: Round values to nearest multiple of 10 (chỉ áp dụng cho các mã hàng không nằm trong danh sách giữ nguyên)
        for row_idx in range(len(processed_df)):
            item_code = str(processed_df.at[row_idx, item_col])

            # Nếu mã hàng không nằm trong danh sách giữ nguyên
            if item_code not in items_to_keep_original:
                # In ra thông tin để debug
                print(f"Áp dụng bước làm tròn cho mã hàng: {item_code}")

                for col_name in numeric_cols:
                    # Apply rounding only to values > 0
                    val = processed_df.at[row_idx, col_name]

                    # Chuyển đổi val thành số nếu có thể
                    try:
                        if pd.isna(val):
                            val = 0
                        else:
                            val = float(val)
                    except (ValueError, TypeError):
                        val = 0

                    if val > 0:
                        rounded_val = int(np.ceil(val / 10.0)) * 10
                        processed_df.at[row_idx, col_name] = rounded_val

                        # In ra thông tin để debug
                        print(f"Đã làm tròn giá trị {val} thành {rounded_val} ở cột {col_name}")

        # Step 4: Khôi phục giá trị gốc cho các mã hàng cần giữ nguyên
        for row_idx in range(len(processed_df)):
            item_code = str(processed_df.at[row_idx, item_col])

            # Nếu mã hàng nằm trong danh sách giữ nguyên
            if item_code in items_to_keep_original:
                for col_idx, col_name in enumerate(processed_df.columns):
                    processed_df.at[row_idx, col_name] = original_df.at[row_idx, col_name]

        # Tạo tên file output
        output_file = os.path.splitext(input_file)[0] + "_output.xlsx"

        # Luôn tạo tên file output mới với số ngẫu nhiên để tránh xung đột
        import random
        random_num = random.randint(1000, 9999)
        output_file = os.path.splitext(input_file)[0] + f"_output_{random_num}.xlsx"
        print(f"Tạo file output mới: {output_file}")

        # Lưu sheet Demand đã xử lý vào file tạm
        temp_demand_file = os.path.splitext(output_file)[0] + "_temp_demand.xlsx"
        processed_df.to_excel(temp_demand_file, sheet_name=demand_sheet_name, index=False)
        print(f"Đã lưu sheet {demand_sheet_name} đã xử lý vào file tạm")

        # Sử dụng phương pháp sao chép file trước, sau đó chỉ thay đổi sheet DEMAND
        print(f"Tạo file output mới bằng cách sao chép file gốc...")

        try:
            # Sao chép file gốc thành file output để giữ nguyên tất cả định dạng và dữ liệu
            import shutil
            shutil.copy2(input_file, output_file)
            print(f"Đã sao chép file gốc thành file output")

            # Sử dụng pandas để chỉ thay thế sheet DEMAND
            # Đọc tất cả các sheet từ file output
            with pd.ExcelFile(output_file) as xls:
                all_dfs = {}
                for sheet in xls.sheet_names:
                    if sheet != demand_sheet_name:
                        # Đọc các sheet khác để giữ nguyên
                        all_dfs[sheet] = pd.read_excel(output_file, sheet_name=sheet)

            # Thêm sheet DEMAND đã xử lý
            demand_df = pd.read_excel(temp_demand_file, sheet_name=demand_sheet_name)
            all_dfs[demand_sheet_name] = demand_df

            # Ghi lại tất cả các sheet vào file output
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
                for sheet_name, df in all_dfs.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    if sheet_name == demand_sheet_name:
                        print(f"Đã ghi sheet {demand_sheet_name} đã xử lý vào file output")
                    else:
                        print(f"Đã giữ nguyên sheet {sheet_name} trong file output")

            print(f"Đã hoàn thành việc tạo file output")

        except Exception as e:
            print(f"Lỗi khi tạo file output: {str(e)}")

            # Thử cách đơn giản hơn nếu cách trên thất bại
            print(f"Thử cách đơn giản hơn...")

            try:
                # Sao chép file gốc thành file output
                import shutil
                shutil.copy2(input_file, output_file)
                print(f"Đã sao chép file gốc thành file output")

                # Chỉ thay thế sheet DEMAND
                with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
                    # Xóa sheet DEMAND cũ nếu tồn tại
                    book = writer.book
                    if demand_sheet_name in book.sheetnames:
                        idx = book.sheetnames.index(demand_sheet_name)
                        book.remove(book.worksheets[idx])
                        print(f"Đã xóa sheet {demand_sheet_name} cũ")

                    # Thêm sheet DEMAND đã xử lý
                    demand_df = pd.read_excel(temp_demand_file, sheet_name=demand_sheet_name)
                    demand_df.to_excel(writer, sheet_name=demand_sheet_name, index=False)
                    print(f"Đã ghi sheet {demand_sheet_name} đã xử lý vào file output")

                print(f"Đã hoàn thành việc tạo file output với cách đơn giản hơn")

            except Exception as e2:
                print(f"Lỗi khi thử cách đơn giản hơn: {str(e2)}")
                print(f"Không thể tạo file output. Vui lòng thử lại.")
                raise e  # Ném lại lỗi ban đầu

        # Xóa file tạm
        try:
            os.remove(temp_demand_file)
        except:
            pass

        print(f"Processing complete. Output saved to: {output_file}")

        return output_file

    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return None

def main():
    # Cho phép người dùng nhập đường dẫn file input
    try:
        input_file = input("Nhập đường dẫn file Excel (nhấn Enter để sử dụng file mặc định 'SV.xlsx'): ")

        # Nếu người dùng không nhập gì, sử dụng file mặc định
        if not input_file.strip():
            input_file = "SV.xlsx"
            print(f"Sử dụng file mặc định: {input_file}")

        # Kiểm tra xem file có tồn tại không
        if not os.path.exists(input_file):
            print(f"Lỗi: File '{input_file}' không tồn tại.")
            return

        # Process the file
        output_file = process_demand_sheet(input_file)

        if output_file:
            print(f"Xử lý dữ liệu thành công. File output: {output_file}")
        else:
            print("Xử lý dữ liệu thất bại.")

    except KeyboardInterrupt:
        print("\nĐã hủy thao tác.")
    except Exception as e:
        print(f"Lỗi: {str(e)}")

if __name__ == "__main__":
    main()
