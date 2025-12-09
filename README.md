# ca-nda-generator

# ca-nda-generator

This project is a web application designed to simplify the generation of Non-Disclosure Agreements (NDAs), specifically tailored for Chartered Accountancy firms. It allows users to fill in various details for Party A (including CA Firm details) and Party B, select Party A information from a pre-defined list, and generate a dynamic NDA based on a customizable template. The filled agreement details are logged for record-keeping.

## Features

*   **User Authentication:** Secure login mechanism (e.g., Google OAuth, as suggested by login/logout functionality).
*   **Dynamic NDA Generation:** Generates NDAs by replacing placeholders in a template file (`sample.txt`) with user-provided data.
*   **Party A Management:**
    *   Select Party A details (Company Name, LLPIN, Address, CA Firm Name, CA Name, Member Reg No, Partner/Proprietor) from a pre-configured Excel sheet (`worked_for.xlsx`).
    *   Option to manually enter/clear Party A details.
*   **Party B Data Input:** Fields for Company Name, CIN, and Address for Party B.
*   **Agreement Details:** Input for the effective date of the agreement.
*   **Data Persistence:** Saves all entered agreement details to an Excel log file (`saved_texts.xlsx`) upon submission.
*   **Preview Functionality:** Allows users to view the generated NDA with filled details before finalization.
*   **Download Options:** (Inferred) Ability to download the generated NDA as a PDF and/or DOCX document.
*   **Customizable Template:** The core NDA text can be easily modified in `sample.txt`.

## Technologies Used

*   **Backend:** Node.js, Express.js
*   **Excel Handling:** `exceljs` library
*   **Authentication:** (Potentially Passport.js with Google Strategy, implied by OAuth flow)
*   **Frontend:** HTML, CSS, JavaScript
*   **PDF Generation:** `jspdf` (client-side)

## Setup and Installation

### Prerequisites

*   Node.js (LTS version recommended)
*   npm (Node Package Manager, usually comes with Node.js)

### Steps

1.  **Clone the Repository:**
    ```bash
    git clone https://github.com/your-username/ca-nda-generator.git
    cd ca-nda-generator
    ```
    (Replace `https://github.com/your-username/ca-nda-generator.git` with your actual repository URL)

2.  **Install Dependencies:**
    ```bash
    npm install
    ```

3.  **Configure Environment Variables:**
    Create a `.env` file in the root directory of your project (if not already present). This file will store sensitive information like OAuth credentials and session secrets.
    ```
    # Example .env content (adjust as per your actual OAuth setup)
    GOOGLE_CLIENT_ID=YOUR_GOOGLE_CLIENT_ID
    GOOGLE_CLIENT_SECRET=YOUR_GOOGLE_CLIENT_SECRET
    SESSION_SECRET=a_strong_random_secret_string_here
    PORT=3000
    ```
    *   Replace `YOUR_GOOGLE_CLIENT_ID` and `YOUR_GOOGLE_CLIENT_SECRET` with your credentials obtained from the Google Cloud Console for OAuth.
    *   Change `SESSION_SECRET` to a long, random string.

4.  **Prepare `worked_for.xlsx`:**
    *   Place an Excel file named `worked_for.xlsx` in the root directory of your project.
    *   This file should contain a sheet (e.g., named "Party A Options") with your pre-defined Party A data.
    *   **Crucially, the header row in `worked_for.xlsx` must match the expected column names used in the application.**
        Expected headers (case-sensitive): `Company Name`, `LLPIN`, `Address`, `CA FIRM NAME`, `CA NAME`, `MEMBER REG NO`, `PARTNER/PROPRIETOR`.

5.  **Customize `sample.txt`:**
    *   Edit the `public/sample.txt` file to define the structure and content of your NDA.
    *   Use `{{PLACEHOLDER_NAME}}` for fields that will be dynamically filled by the application.
    *   **Ensure placeholders exactly match the names expected by the application.**
        Examples: `{{Date}}`, `{{Company Name A}}`, `{{LLPIN}}`, `{{Address A}}`, `{{CA FIRM NAME}}`, `{{CA NAME}}`, `{{MEMBER REG NO}}`, `{{PARTNER/PROPRIETOR}}`, `{{Company Name B}}`, `{{CIN}}`, `{{Address B}}`.

### Running the Application

1.  **Start the Server:**
    ```bash
    node index.js
    ```
    The server will typically run on `http://localhost:3000` (or the PORT defined in your `.env` file).

2.  **Access the Application:**
    Open your web browser and navigate to `http://localhost:3000`.

## Usage

1.  **Login:** You will be redirected to a login page (e.g., Google OAuth). Authenticate to proceed.
2.  **Fill Agreement Details:** On the `profile.html` page:
    *   Enter the `Date`.
    *   For **Party A Company Details**, you can either:
        *   Click "Select from List" to choose from entries in `worked_for.xlsx`.
        *   Manually type in the details.
    *   Fill in **Party B Details** (Company Name B, CIN, Address B).
    *   The **CA Firm Details for Party A** section will automatically populate if you select Party A from the list.
3.  **Save Agreement:** Click the "Save Agreement Details" button. This will save all current inputs to `saved_texts.xlsx`.
4.  **View Last Saved Entry:** Click "View Last Saved Entry" to see the dynamically generated NDA on `display.html`.
5.  **Download:** (If implemented) Use the "Download PDF" or "Download DOC" buttons on `display.html` to save the agreement.

## Important Notes on Data Integrity

To avoid swapped or missing values in the generated NDA:

1.  **`sample.txt` Placeholders:** Ensure that `public/sample.txt` has the exact placeholders (e.g., `{{CIN}}`, `{{CA NAME}}`) in the correct positions where you want their respective values to appear.
2.  **`worked_for.xlsx` Headers:** The header row in `worked_for.xlsx` (for Party A options) must have the exact column names: `Company Name`, `LLPIN`, `Address`, `CA FIRM NAME`, `CA NAME`, `MEMBER REG NO`, `PARTNER/PROPRIETOR`.
3.  **`saved_texts.xlsx` Headers and `index.js` `addRow` Order:**
    *   The header row in `saved_texts.xlsx` must match the order in which data is sent to the `worksheet.addRow()` function in `index.js`.
    *   The `worksheet.addRow([...])` call in `index.js` (within the `/api/save-text` endpoint) must pass variables in an order that precisely corresponds to the column headers in `saved_texts.xlsx`.

If values are still swapped after reviewing `sample.txt`, the issue is almost certainly in the `worksheet.addRow()` order in `index.js` not matching the `saved_texts.xlsx` column headers. Restarting the Node.js server and re-saving data after any `index.js` changes is crucial.
