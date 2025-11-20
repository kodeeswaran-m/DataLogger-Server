const mongoose = require("mongoose");

const CallSchema = new mongoose.Schema(
  {
    checked: { type: Boolean, default: false },
    notes: { type: String, trim: true, default: "" },
  },
  { _id: false }
);

const ProspectDetailSchema = new mongoose.Schema({
  month: { type: String, trim: true, required: true },
  quarter: { type: String, trim: true, required: true },
  prospect: { type: String, trim: true, required: true },
  geo: { type: String, trim: true },
  lob: { type: String, trim: true },
  call1: { type: CallSchema, default: () => ({}) },
  call2: { type: CallSchema, default: () => ({}) },
  call3: { type: CallSchema, default: () => ({}) },
  coreOfferings: { type: String, trim: true },
  primaryNeed: { type: String, trim: true },
  secondaryNeed: { type: String, trim: true },
  category: { type: String, trim: true },
  trace: { type: String, trim: true },
  salesSpoc: { type: String, trim: true },
  oppId: { type: String, trim: true, index: true },
  oppDetails: { type: String, trim: true },
  deck: { type: String, trim: true }, // could be URL or filename
  deckPublicId: { type: String, trim: true }, // ADD THIS
  rag: { type: String, trim: true }, // RAG status (e.g., Red/Amber/Green)
  remark: { type: String, trim: true },

  // metadata
  createdAt: { type: Date, default: Date.now },
  updatedAt: { type: Date, default: Date.now },
});

// Update updatedAt on save/update
ProspectDetailSchema.pre("save", function (next) {
  this.updatedAt = Date.now();
  next();
});

module.exports = mongoose.model("ProspectDetail", ProspectDetailSchema);
