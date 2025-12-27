## Installation & Setup
These instructions assume a fresh installation on a new machine.

### Prerequisites
- **Node.js**: Version 18 or higher.
- **Office 365**: Word Desktop or Word on the Web.
- **Gemini API Key**: Get one from [Google AI Studio](https://aistudio.google.com/app/apikey). For the sake of this challenge, please contact me directly if you need a key.

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
1. Create a file named `env.local` in the `apps/web-client` directory.
2. Add your API key to it:
   ```env
   GEMINI_API_KEY=your_api_key_here
   ```
3. Rename it to `.env.local` (the `.` is important). You might need the terminal to do this, since the .env is often reserved for system files.
   - On mac, navigate to the `apps/web-client` directory and run `mv env.local .env.local` to rename it. The file should disappear from the folder.
   - On windows, navigate to the `apps/web-client` directory and run `ren env.local .env.local` to rename it. The file should disappear from the folder.

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

