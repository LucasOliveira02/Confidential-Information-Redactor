import { createServer } from 'https';
import { parse } from 'url';
import next from 'next';
import * as devCerts from "office-addin-dev-certs";

const dev = process.env.NODE_ENV !== 'production';
const hostname = 'localhost';
const port = 3000;

// When using middleware `hostname` and `port` must be provided below
const app = next({ dev, hostname, port });
const handle = app.getRequestHandler();

app.prepare().then(async () => {
    try {
        console.log("Loading development certificates...");
        const options = await devCerts.getHttpsServerOptions();

        createServer(options, async (req, res) => {
            try {
                // Be sure to pass `true` as the second argument to `url.parse`.
                // This tells it to parse the query portion of the URL.
                const parsedUrl = parse(req.url || '', true);

                await handle(req, res, parsedUrl);
            } catch (err) {
                console.error('Error occurred handling', req.url, err);
                res.statusCode = 500;
                res.end('internal server error');
            }
        })
            .once('error', (err) => {
                console.error(err);
                process.exit(1);
            })
            .listen(port, () => {
                console.log(`> Ready on https://${hostname}:${port}`);
            });

    } catch (e) {
        console.error("Failed to set up HTTPS:", e);
        process.exit(1);
    }
});
