const mongoose = require("mongoose");

mongoose.connect("mongodb://localhost:27017/panone-limited", {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

const TripSchema = new mongoose.Schema(
  {},
  {
    timestamps: true,
    strict: false,
  }
);

module.exports = mongoose.model("trips", TripSchema);
