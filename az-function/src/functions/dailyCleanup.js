const { app } = require("@azure/functions");
const { runCleanup } = require("../../graphClient");

app.timer("dailyCleanup", {
  schedule: "0 0 */4 * * *", // every 4 hours
  handler: async (myTimer, context) => {
    try {
      const result = await runCleanup();
      context.log(
        `Scanned: ${result.scanned}. Matched: ${result.matched}. Deleted: ${result.deleted}.`
      );
    } catch (error) {
      context.error(`Cleanup failed: ${error.message}`);
      throw error;
    }
  },
});
