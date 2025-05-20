# VBA 巨集 Excel 資料處理工具

本儲存庫包含一組 VBA 巨集，專為處理和整理 Excel 檔案中的財務資料而設計，用於資料對帳和報表生成。每個巨集針對特定財務資料類型（如帳變、支付渠道、交易記錄等）提供專屬功能。以下是各巨集的功能概述、使用說明和檔案結構。

## 目錄
- [巴西帳變整理](#巴西帳變整理)
- [BC支付渠道整理](#bc支付渠道整理)
- [GOSM整理](#gosm整理)
- [菲包充提整理](#菲包充提整理)
- [非菲充提整理](#非菲充提整理)
- [使用說明](#使用說明)
- [檔案結構](#檔案結構)
- [系統需求](#系統需求)
- [注意事項](#注意事項)
- [English Version](#english-version)

## 巴西帳變整理
**目的**：將多個包含帳變資料的 Excel 檔案合併到單一目標工作簿，並進行格式化和資料清理。

**功能**：
- 提示使用者選擇多個 Excel 檔案（.xlsx、.xls）進行合併。
- 將資料合併到新工作簿中，後續檔案的資料從第二行開始附加（排除標題）。
- 套用格式：設置字型為 Calibri（大小 11），使用固定寬度文字分列，移除多餘欄位，並清理特定欄位（例如移除 G 欄中的 "-*"）。
- 將合併後的資料儲存到使用者選擇的目標 Excel 檔案，並刷新其中的樞紐分析表。
- 允許連續合併，提示使用者是否處理更多檔案。

**主要特色**：
- 支援多檔案連續處理。
- 統一格式化與資料清理。
- 保留目標檔案結構並更新樞紐分析表。

## BC支付渠道整理
**目的**：將多個支付渠道資料的 Excel 檔案合併並整理成單一格式化工作簿。

**功能**：
- 提示使用者選擇多個 Excel 檔案（.xlsx、.xls）。
- 建立新工作簿，將每個檔案的資料複製並新增「平台」欄位（基於檔案名前綴）。
- 套用格式：調整日期格式、分列、插入計算欄位（例如從時間提取小時）、篩選「success」狀態，並自動調整欄寬。
- 以標準化名稱（例如「MMDD-九平台支付.xlsx」）儲存合併後的工作簿，日期為前一天。

**主要特色**：
- 自動新增平台識別欄位。
- 篩選成功交易資料。
- 自動化日期調整與欄位格式化。

## GOSM整理
**目的**：處理並合併 GOSM 相關財務資料到主工作簿。

**功能**：
- 提示使用者選擇包含主 GOSM 工作簿（檔名含「GOSM-」）和相關資料檔案（例如「代收」、「代付」、「提出」、「入」）的資料夾。
- 處理資料檔案：
  - 將日期時間欄位拆分為日期和時間。
  - 刪除多餘欄位並格式化日期。
  - 根據檔案類型將資料合併到主工作簿的第一或第二工作表。
- 自動填充指定欄位的公式並刷新主工作簿中的樞紐分析表。
- 報告遺漏的檔案（若有）並儲存更新後的主工作簿。

**主要特色**：
- 支援多種檔案類型並進行特定處理。
- 自動化公式填充與樞紐分析表更新。
- 提供遺漏檔案的回饋。

## 菲包充提整理
**目的**：處理菲律賓相關的交易（TR）檔案並合併到 AP 模板工作簿。

**功能**：
- 提示使用者選擇包含子資料夾（內含「TR_*.xlsx」檔案）和 AP 模板（.xlsm）的資料夾。
- 根據子資料夾名稱重新命名 TR 檔案並清理資料（例如將「+」替換為「充值」，「-」替換為「提款」，分列，格式化日期）。
- 將清理後的 TR 資料複製到 AP 模板的「上下分紀錄」工作表，從第二行開始貼上。
- 以名稱如「(資料夾名稱)MMDD充提.xlsm」儲存更新後的 AP 工作簿到外層資料夾。

**主要特色**：
- 遞迴處理子資料夾。
- 使用模板確保輸出一致性。
- 自動化資料清理與複製。

## 非菲充提整理
**目的**：處理非菲律賓相關的交易（TR）檔案並合併到現有的 AP 工作簿。

**功能**：
- 提示使用者選擇包含子資料夾（內含「TR_*.xlsx」檔案）的資料夾。
- 根據子資料夾名稱重新命名 TR 檔案並清理資料（與菲包充提整理類似）。
- 在每個子資料夾中識別 AP 檔案（包含「(AP」、「(TG」、「(US」等標籤），並將 TR 資料複製到其「上下分紀錄」工作表。
- 根據最新交易日期更新 AP 工作簿中的樞紐分析表並儲存。

**主要特色**：
- 支援多種 AP 檔案類型。
- 動態更新樞紐分析表。
- 處理基於子資料夾的資料。

## 使用說明
1. **啟用巨集**：確保 Excel 已啟用巨集以執行 VBA 腳本。
2. **準備檔案**：將輸入檔案放置在各巨集所需的資料夾結構中（例如 TR 檔案的子資料夾、GOSM 主工作簿或 AP 模板）。
3. **執行巨集**：
   - 開啟包含這些巨集的 Excel 檔案。
   - 從 VBA 編輯器執行所需巨集，或將其分配到按鈕。
   - 根據提示選擇輸入檔案或資料夾。
4. **檢查輸出**：在指定位置檢查生成的或更新的 Excel 檔案。
5. **重複執行（若適用）**：部分巨集（例如巴西帳變整理）允許透過提示連續處理更多檔案。

## 檔案結構
- **輸入檔案**：
  - Excel 檔案（.xlsx、.xls、.xlsm），需符合特定命名規則（例如「TR_*.xlsx」、「GOSM-*.xlsx」、包含「AP」、「代收」等）。
  - 按各巨集要求組織在資料夾或子資料夾中。
- **輸出檔案**：
  - 合併後的 Excel 檔案（.xlsx 或 .xlsm），儲存到使用者指定或預定義的位置。
  - 示例輸出名稱：「MMDD-九平台支付.xlsx」、「(資料夾名稱)MMDD充提.xlsm」或更新的主工作簿。

## 系統需求
- **Microsoft Excel**：支援 VBA 的版本（例如 Excel 2016 或更新版本）。
- **Windows 作業系統**：巨集使用 `Scripting.FileSystemObject` 進行檔案處理，僅適用於 Windows。
- **檔案存取權限**：確保對輸入和輸出檔案目錄具有讀寫權限。
- **巨集啟用環境**：儲存包含巨集的工作簿為 .xlsm 格式。

## 注意事項
- **錯誤處理**：巨集包含錯誤處理機制，以應對檔案存取問題或遺漏檔案，並透過訊息框通知使用者。
- **效能**：執行期間會停用螢幕更新和提示以提升效能，完成後會重新啟用。
- **檔案命名**：確保輸入檔案遵循預期的命名規則，以避免處理錯誤。
- **備份**：執行巨集前請備份檔案，因為巨集可能會覆蓋或修改現有資料。
- **除錯**：如遇問題，可檢查 VBA 除錯輸出（Debug.Print）以獲取詳細處理日誌。

如有問題或貢獻，請在此儲存庫開啟 issue 或 pull request。

---

## English Version

# VBA Macros for Excel Data Processing

This repository contains a collection of VBA macros designed to process and organize Excel files for financial data reconciliation and reporting. Each macro serves a specific purpose, handling different types of financial data such as account changes, payment channels, and transaction records. Below is an overview of each macro, its functionality, and usage instructions.

## Table of Contents
- [Brazil Account Change Consolidation](#brazil-account-change-consolidation)
- [BC Payment Channel Consolidation](#bc-payment-channel-consolidation)
- [GOSM Data Consolidation](#gosm-data-consolidation)
- [Philippines Transaction Consolidation](#philippines-transaction-consolidation)
- [Non-Philippines Transaction Consolidation](#non-philippines-transaction-consolidation)
- [Usage Instructions](#usage-instructions)
- [File Structure](#file-structure)
- [Requirements](#requirements)
- [Notes](#notes)

## Brazil Account Change Consolidation
**Purpose**: Consolidates multiple Excel files containing account change data into a single target workbook, with formatting and data cleaning.

**Functionality**:
- Prompts the user to select multiple Excel files (.xlsx, .xls) to merge.
- Combines data into a new workbook, appending rows from subsequent files (excluding headers after the first file).
- Applies formatting: sets font to Calibri (size 11), splits columns using fixed-width text-to-columns, removes unnecessary columns, and cleans specific columns (e.g., removing "-*" from column G).
- Saves the merged data into a user-selected target Excel file and refreshes any pivot tables.
- Allows continuous merging by prompting the user to process additional files.

**Key Features**:
- Handles multiple files in a loop.
- Formats and cleans data consistently.
- Preserves the target file's structure and refreshes pivot tables.

## BC Payment Channel Consolidation
**Purpose**: Merges and processes payment channel data from multiple Excel files into a single formatted workbook.

**Functionality**:
- Prompts the user to select multiple Excel files (.xlsx, .xls).
- Creates a new workbook and copies data from each file, adding a "Platform" column based on the file name prefix.
- Applies formatting: adjusts dates, splits columns, inserts calculated columns (e.g., extracting hour from time), filters for "success" status, and autofits columns.
- Saves the consolidated workbook with a standardized name (e.g., "MMDD-九平台支付.xlsx") based on the previous day's date.

**Key Features**:
- Adds platform identifiers to data.
- Filters for successful transactions.
- Automates date adjustments and column formatting.

## GOSM Data Consolidation
**Purpose**: Processes and consolidates GOSM-related financial data from multiple files into a main workbook.

**Functionality**:
- Prompts the user to select a folder containing a main GOSM workbook (with "GOSM-" in the name) and related data files (e.g., "代收", "代付", "提出", "入").
- Processes data files:
  - Splits date-time columns into separate date and time columns.
  - Deletes unnecessary columns and formats dates.
  - Merges data into the main workbook's first and second sheets based on file type.
- Auto-fills formulas in specified columns and refreshes pivot tables in the main workbook.
- Reports missing files (if any) and saves the updated main workbook.

**Key Features**:
- Handles multiple file types with specific processing rules.
- Automates formula filling and pivot table refreshing.
- Provides feedback on missing files.

## Philippines Transaction Consolidation
**Purpose**: Processes transaction (TR) files in subfolders and consolidates them into an AP template workbook for Philippines-specific data.

**Functionality**:
- Prompts the user to select a main folder containing subfolders with TR files ("TR_*.xlsx") and an AP template (.xlsm).
- Renames TR files based on the subfolder name and cleans their data (e.g., replaces "+" with "充值", "-" with "提款", splits columns, formats dates).
- Copies cleaned TR data into the AP template's "上下分紀錄" sheet, starting from row 2.
- Saves the updated AP workbook in the main folder with a name like "(FolderName)MMDD充提.xlsm".

**Key Features**:
- Processes subfolders recursively.
- Uses a template-based approach for consistent output.
- Automates data cleaning and copying.

## Non-Philippines Transaction Consolidation
**Purpose**: Processes transaction (TR) files in subfolders and consolidates them into an existing AP workbook for non-Philippines data.

**Functionality**:
- Prompts the user to select a main folder containing subfolders with TR files ("TR_*.xlsx").
- Renames TR files based on the subfolder name and cleans their data (similar to Philippines Transaction Consolidation).
- Identifies an AP file in each subfolder (based on tags like "(AP", "(TG", "(US", etc.) and copies TR data into its "上下分紀錄" sheet.
- Updates pivot tables in the AP workbook based on the latest transaction date and saves the workbook.

**Key Features**:
- Supports multiple AP file types.
- Updates pivot tables dynamically.
- Handles subfolder-based processing.

## Usage Instructions
1. **Enable Macros**: Ensure macros are enabled in Excel to run the VBA scripts.
2. **Prepare Files**: Place input files in the appropriate folder structure as required by each macro (e.g., subfolders for TR files, a main GOSM workbook, or AP templates).
3. **Run Macros**:
   - Open the Excel file containing these macros.
   - Run the desired macro from the VBA editor or assign it to a button.
   - Follow the file/folder selection prompts to choose input files or directories.
4. **Review Output**: Check the generated or updated Excel files in the specified locations for results.
5. **Repeat (if applicable)**: Some macros (e.g., Brazil Account Change Consolidation) allow continuous processing by prompting for additional files.

## File Structure
- **Input Files**:
  - Excel files (.xlsx, .xls, .xlsm) with specific naming conventions (e.g., "TR_*.xlsx", "GOSM-*.xlsx", files containing "AP", "代收", etc.).
  - Organized in folders or subfolders as required by each macro.
- **Output Files**:
  - Consolidated Excel files (.xlsx or .xlsm) saved in user-specified or predefined locations.
  - Example output names: "MMDD-九平台支付.xlsx", "(FolderName)MMDD充提.xlsm", or updated main workbooks.

## Requirements
- **Microsoft Excel**: Version supporting VBA (e.g., Excel 2016 or later).
- **Windows OS**: Macros use `Scripting.FileSystemObject` for file handling, which is Windows-specific.
- **File Access**: Ensure read/write access to input and output file directories.
- **Macro-Enabled Environment**: Save the workbook containing these macros as .xlsm.

## Notes
- **Error Handling**: Macros include error handling to manage file access issues or missing files, with user notifications via message boxes.
- **Performance**: Screen updating and alerts are disabled during execution to improve performance; they are re-enabled upon completion.
- **File Naming**: Ensure input files follow expected naming conventions to avoid processing errors.
- **Backup**: Always back up files before running macros, as they may overwrite or modify existing data.
- **Debugging**: Check the VBA debug output (Debug.Print) for detailed processing logs if issues arise.

For issues or contributions, please open an issue or pull request on this repository.