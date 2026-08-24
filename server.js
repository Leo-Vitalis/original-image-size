const https = require("https");
const fs = require("fs");
const path = require("path");

const devCerts =
    require("office-addin-dev-certs");


(async () => {

    const options =
        await devCerts.getHttpsServerOptions();


    const server = https.createServer(
        options,
        function (req, res) {

            let url =
                decodeURIComponent(
                    req.url.split("?")[0]
                );


            if (url === "/") {

                url = "/taskpane.html";

            }


            const filePath =
                path.join(__dirname, url);


            fs.readFile(
                filePath,
                function (error, data) {

                    if (error) {

                        res.writeHead(404);

                        res.end("Not found");

                        return;

                    }


                    const extension =
                        path.extname(filePath)
                        .toLowerCase();


                    const contentTypes = {

                        ".html":
                            "text/html; charset=utf-8",

                        ".js":
                            "text/javascript; charset=utf-8",

                        ".css":
                            "text/css; charset=utf-8",

                        ".png":
                            "image/png",

                        ".xml":
                            "application/xml"

                    };


                    res.writeHead(
                        200,
                        {
                            "Content-Type":
                                contentTypes[extension]
                                || "application/octet-stream",

                            "Access-Control-Allow-Origin":
                                "*"
                        }
                    );


                    res.end(data);

                }
            );

        }
    );


    server.listen(
        3000,
        "localhost",
        function () {

            console.log(
                "Original Image Size add-in running at:"
            );

            console.log(
                "https://localhost:3000"
            );

            console.log(
                "Keep this window open while using the add-in."
            );

        }
    );

})();