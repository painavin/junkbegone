// Run this once, locally, to sign in via device code and seed the blob-backed
// token cache. After this, the Azure Function refreshes silently on its own.
require("dotenv").config();

const { createPca, getContainerClient, SENDERS_BLOB, CONTAINER_NAME } = require("./graphClient");

async function ensureSendersBlobExists(containerClient) {
  const blockBlob = containerClient.getBlockBlobClient(SENDERS_BLOB);
  if (await blockBlob.exists()) return;

  const placeholder = JSON.stringify(["placeholder@example.com"], null, 2);
  await blockBlob.upload(placeholder, Buffer.byteLength(placeholder));
  console.log(
    `Created placeholder ${SENDERS_BLOB} in container "${CONTAINER_NAME}" — edit it with your real conservative sender list before relying on the schedule.`
  );
}

async function main() {
  const containerClient = getContainerClient();
  await containerClient.createIfNotExists();
  await ensureSendersBlobExists(containerClient);

  const pca = createPca(containerClient);

  const result = await pca.acquireTokenByDeviceCode({
    scopes: ["Mail.ReadWrite", "User.Read"],
    deviceCodeCallback: (response) => {
      console.log(response.message);
    },
  });

  console.log(`Signed in as ${result.account.username}. Token cache saved to blob storage.`);
}

main().catch((err) => {
  console.error("Bootstrap login failed:", err);
  process.exit(1);
});
