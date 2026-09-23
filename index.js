const express = require("express");
const app = express();

const PORT = process.env.PORT || 3021;

app.get("/", (req, res) => {
  res.send("ST-BOT is running!");
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Server running on port ${PORT}`);
});
