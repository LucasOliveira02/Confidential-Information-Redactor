import * as fs from "fs";
import * as path from "path";
import { getHttpsServerOptions } from "office-addin-dev-certs";

async function verifyCerts() {
    try {
        console.log("Verifying developer certificates...");
        const options = await getHttpsServerOptions();

        if (options.ca && options.key && options.cert) {
            fs.writeFileSync(path.resolve(process.cwd(), "cert.pem"), options.cert);
            fs.writeFileSync(path.resolve(process.cwd(), "key.pem"), options.key);
            console.log("Certificates verified and copied to project root.");
        } else {
            throw new Error("Unable to retrieve certificate options.");
        }
    } catch (err) {
        console.error("Error verifying certificates:", err);
        process.exit(1);
    }
}

verifyCerts();
