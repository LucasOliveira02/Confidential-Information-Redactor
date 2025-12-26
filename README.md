# Challenge: Document Redaction with Tracking & Confidentiality Header

Create a button that will redact sensitive information from a Word document, add a confidentiality header, and enable the Tracking Changes feature to log these modifications.

## Requirements:

1. **Redact Sensitive Information**
    - Retrieve the document's complete content
    - Locate and identify sensitive information (emails, phone numbers, social security numbers)
    - Replace this information with redaction markers in the document
2. **Add Confidential Header**
    - Insert a header at the top of the document stating "CONFIDENTIAL DOCUMENT"
    - Ensure this header addition is tracked by the Tracking Changes feature
3. **Enable Tracking Changes**
    - Use the Office Tracking Changes API to enable tracking changes
    - Make sure to only use Tracking Changes if the Word API is available
    [Word JavaScript API requirement set 1.5 - Office Add-ins | Microsoft Learn](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-1-5-requirement-set?view=word-js-preview)

## Technical Requirements

- Code must be with TypeScript.
- You are free to use any framework or build tool (Vite, Next.js, etc.),  DON'T use the minimum setup of this repo.
- Use self written CSS for styling instead of external libraries, we expect good design and craftsmanship.
- The solution must run in Word on the web or Word desktop.
- Don't create a public forked repository, otherwise your solution will be disqualified. Share your solution as a private repository or a zip file.
- One of the evaluation criteria will be code quality, so please ensure your code is clean, well-structured, and follows best practices and crasftsmanship (non AI).


## Testing Your Solution
Use the attached Document-To-Be-Redacted.docx file to test your solution. The document contains various instances of sensitive information that should be redacted when your add-in is executed.

We will use a different document to evaluate your solution, so ensure that your redaction logic is robust and can handle various scenarios.

## Submission

1. Ensure your solution meets all the requirements outlined above.
2. Share your solution as a zip file (without the node_modules folder).
3. Include any necessary instructions to run and test your solution, BUT it should be straightforward to run following the steps in the "Run the Challenge" section.
4. Submit your solution to yotam.segal@mccarren.ai before the deadline specified in the challenge announcement.

Good luck, and we look forward to seeing your innovative solutions!

## Installation & Setup
These instructions assume a fresh installation on a new machine.

### Prerequisites
- **Node.js**: Version 18 or higher.
- **Office 365**: Word Desktop or Word on the Web.
- **Gemini API Key**: Get one from [Google AI Studio](https://aistudio.google.com/app/apikey). For the sake of the interview please contact me directly if you need a key.

### 1. Setup the Project
1. Unzip the project or clone the repository.
2. Navigate to the web-client folder:
   ```bash
   cd apps/web-client
   ```
3. Install dependencies:
   ```bash
   npm install
   ```

### 2. Configure Environment Secrets
This project requires a Gemini API key to function.
1. Create a file named `.env.local` in the `apps/web-client` directory.
2. Add your API key to it:
   ```env
   GEMINI_API_KEY=your_api_key_here
   ```

### 3. Run the Development Server
1. Start the server and sideload the add-in:
   ```bash
   npm run dev
   ```
   *This runs the server at `https://localhost:3000`*

### 4. Load into Word
**Desktop (Mac/Windows):**
- The `npm run dev` command might attempt to sideload automatically.
- If it doesn't appear, go to **Home > Add-ins > Add headers** and the side panel should appear.

**Word on the Web:**
- Go to [Word Online](https://word.office.com/).
- Create a new document.
- Go to **Home > Add-ins > More Add-ins > My Add-ins > Upload My Add-in**.
- Select the `manifest.xml` file located in `apps/web-client` and load it.
- Go to **Home > Show Taskpane**.
- The side panel should appear.

### 5. Trust Development Certificates (Crucial)
If you see a blank white panel or connection errors:
1. Open `https://localhost:3000` in your browser.
2. If your browser warns "Your connection is not private", click **Advanced > Proceed to localhost (unsafe)**.
3. Once you see the Redactor page, the certificate is trusted for this session.
4. Refresh the task pane in Word.

