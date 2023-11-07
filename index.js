// Import required modules
const express = require("express");
const app = express();
const path = require("path");
const { authenticate } = require("@google-cloud/local-auth");
const { google } = require("googleapis");
const rateLimit = require("express-rate-limit");

const port = 5000;

// OAuth2 scopes required for Gmail access
const SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.labels",
    "https://mail.google.com/",
];

// Gmail label used for Vacation Responder
const labelName = "Vacation Responder";


// rate limiter middleware
const limiter = rateLimit({
    windowMs: 60 * 1000, // 1 minute
    max: 10, // 10 requests per minute
});

app.use(limiter);

// function to create a Gmail label
async function createLabel(gmail) {
    try {

        const response = await gmail.users.labels.create({
            userId: "me",
            requestBody: {
                name: labelName,
                labelListVisibility: "labelShow",
                messageListVisibility: "show",
            },
        });
        // Return the ID of the newly created or existing label
        return response.data.id;
    } catch (error) {
        if (error.code === 409) {
            // If the label already exists, find and return its ID
            const response = await gmail.users.labels.list({
                userId: "me",
            });
            const label = response.data.labels.find((label) => label.name === labelName);
            return label.id;
        } else {

            throw error;
        }
    }
}

// function to get unread messages from the inbox
async function getUnreadMessages(gmail) {
    // Retrieve unread messages from the inbox
    const response = await gmail.users.messages.list({
        userId: "me",
        labelIds: ["INBOX"],
        q: "is:unread",
    });
    // Return the list of unread messages or an empty array
    return response.data.messages || [];
}

// function to check if a message has been replied to
function hasReplied(email) {

    return email.payload.headers.some((header) => header.name === "In-Reply-To");
}

// function to send a vacation response
async function sendVacationResponse(gmail, email, labelId) {

    const replyMessage = {
        userId: "me",
        resource: {
            raw: Buffer.from(
                `To: ${email.payload.headers.find((header) => header.name === "From").value}\r\n` +
                `Subject: Re: ${email.payload.headers.find((header) => header.name === "Subject").value}\r\n` +
                `Content-Type: text/plain; charset="UTF-8"\r\n` +
                `Content-Transfer-Encoding: 7bit\r\n\r\n` +
                `Thank you for your message. I'm currently out of the office and will get back to you when I return.\r\n`
            ).toString("base64"),
        },
    };
    // Send the response and update labels
    await gmail.users.messages.send(replyMessage);
    await gmail.users.messages.modify({
        userId: "me",
        id: email.id,
        resource: {
            addLabelIds: [labelId],
            removeLabelIds: ["INBOX"],
        },
    });
}

async function main() {

    try {
        // Authenticate using the provided credentials and OAuth2 scopes
        const auth = await authenticate({
            keyfilePath: path.join(__dirname, "credentials.json"),
            scopes: SCOPES,
        });

        // Initialize the Gmail API client
        const gmail = google.gmail({ version: "v1", auth });

        // Get or create the Vacation Responder label and obtain its ID
        const labelId = await createLabel(gmail);

        // Set an interval to periodically check for unread messages and send responses
        setInterval(async () => {
            const messages = await getUnreadMessages(gmail);

            if (messages && messages.length > 0) {
                for (const message of messages) {
                    const messageData = await gmail.users.messages.get({
                        userId: "me",
                        id: message.id,
                    });

                    const email = messageData.data;

                    if (!hasReplied(email)) {
                        await sendVacationResponse(gmail, email, labelId);
                    }
                }
            }
        }, Math.floor(Math.random() * (120 - 45 + 1) + 45) * 1000);

        // Start the Express server
        app.listen(port, () => {
            console.log(`Server is running on port ${port}`);
        });
    } catch (error) {
        // Handle any errors during authentication or server startup
        console.error(`Failed to process request: ${error}`);
        app.use((err, req, res, next) => {
            console.error(err.stack);
            res.status(500).send('Something broke!');
        });
    }
}

main();