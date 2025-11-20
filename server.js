// server.js
require("dotenv").config();
const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const prospectRoutes = require("./routes/prospectDetailRoutes");
const morgan = require("morgan");
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(morgan("dev"))
app.use(express.json()); // parse JSON
app.use(express.urlencoded({ extended: true })); // parse urlencoded

// Routes
app.use("/api/prospectDetail", prospectRoutes);

// MongoDB connection
mongoose
  .connect(process.env.MONGO_URI, {
    // no need for useNewUrlParser / useUnifiedTopology in mongoose v6+
  })
  .then(() => console.log("MongoDB connected"))
  .catch((err) => console.error("MongoDB connection error:", err));

// Start server
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
