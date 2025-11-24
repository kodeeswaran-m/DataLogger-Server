// controllers/prospectDetailController.js
const ProspectDetail = require("../models/prospectDetailModel");
const cloudinary = require("../config/cloudinary");
const fs = require("fs");
const ExcelJS = require("exceljs");

const safeParseJSON = (maybe) => {
  if (!maybe) return null;
  if (typeof maybe === "object") return maybe;
  try {
    return JSON.parse(maybe);
  } catch (e) {
    return null;
  }
};

// Always force single string (defensive against FormData arrays)
const normalizeField = (v) => {
  if (v === undefined || v === null) return "";
  if (Array.isArray(v)) return String(v[0] || "");
  return String(v);
};

exports.createProspectDetail = async (req, res) => {
  try {
    let deckUrl = "";
    let deckPublicId = "";

    if (req.file) {
      const result = await cloudinary.uploader.upload(req.file.path, {
        folder: "decks",
        resource_type: "raw",
      });
      deckUrl = result.secure_url;
      deckPublicId = result.public_id;
      try {
        fs.unlinkSync(req.file.path);
      } catch (e) {}
    }

    const body = req.body || {};

    const call1 = safeParseJSON(body.call1) || {
      checked: Boolean(
        body.call1_checked === "true" || body.call1_checked === true
      ),
      notes: body.call1_notes || "",
    };
    const call2 = safeParseJSON(body.call2) || {
      checked: Boolean(
        body.call2_checked === "true" || body.call2_checked === true
      ),
      notes: body.call2_notes || "",
    };
    const call3 = safeParseJSON(body.call3) || {
      checked: Boolean(
        body.call3_checked === "true" || body.call3_checked === true
      ),
      notes: body.call3_notes || "",
    };

    // normalize category inputs (force single string)
    const categoryRaw = normalizeField(body.category);
    const categoryOtherRaw = normalizeField(body.categoryOther);

    // decide saved category (if dropdown = 'other' then use categoryOther)
    const savedCategory =
      categoryRaw === "other" ? categoryOtherRaw || "" : categoryRaw || "";
    const savedCategoryOther =
      categoryRaw === "other" ? categoryOtherRaw || "" : "";

    const doc = new ProspectDetail({
      month: normalizeField(body.month),
      quarter: normalizeField(body.quarter),
      prospect: normalizeField(body.prospect),
      geo: normalizeField(body.geo),
      lob: normalizeField(body.lob),

      call1,
      call2,
      call3,

      coreOfferings: normalizeField(body.coreOfferings),
      primaryNeed: normalizeField(body.primaryNeed),
      secondaryNeed: normalizeField(body.secondaryNeed),

      category: savedCategory,
      categoryOther: savedCategoryOther,

      trace: normalizeField(body.trace),
      salesSpoc: normalizeField(body.salesSpoc),
      oppId: body.oppId || `OPP-${Date.now()}`,
      oppDetails: normalizeField(body.oppDetails),

      deck: deckUrl,
      deckPublicId,
      rag: normalizeField(body.rag),
      remark: normalizeField(body.remark),
    });

    await doc.save();
    return res.status(201).json({ success: true, data: doc });
  } catch (err) {
    console.error("createProspectDetail error:", err);
    return res.status(500).json({ success: false, message: "Server error" });
  }
};

exports.getProspectById = async (req, res) => {
  try {
    const doc = await ProspectDetail.findById(req.params.id);
    if (!doc)
      return res.status(404).json({ success: false, message: "Not found" });
    return res.json({ success: true, data: doc });
  } catch (err) {
    console.error("getProspectById error:", err);
    return res.status(500).json({ success: false, message: "Server error" });
  }
};

exports.updateProspect = async (req, res) => {
  try {
    const id = req.params.id;
    const existing = await ProspectDetail.findById(id);
    if (!existing)
      return res.status(404).json({ success: false, message: "Not found" });

    const body = req.body || {};
    let deckUrl = existing.deck;
    let deckPublicId = existing.deckPublicId;

    if (req.file) {
      if (deckPublicId) {
        try {
          await cloudinary.uploader.destroy(deckPublicId, {
            resource_type: "raw",
          });
        } catch (e) {
          /*ignore*/
        }
      }
      const result = await cloudinary.uploader.upload(req.file.path, {
        folder: "decks",
        resource_type: "raw",
      });
      deckUrl = result.secure_url;
      deckPublicId = result.public_id;
      try {
        fs.unlinkSync(req.file.path);
      } catch (e) {}
    }

    const call1 = safeParseJSON(body.call1) || {
      checked: Boolean(
        body.call1_checked === "true" || body.call1_checked === true
      ),
      notes: body.call1_notes ?? existing.call1?.notes ?? "",
    };
    const call2 = safeParseJSON(body.call2) || {
      checked: Boolean(
        body.call2_checked === "true" || body.call2_checked === true
      ),
      notes: body.call2_notes ?? existing.call2?.notes ?? "",
    };
    const call3 = safeParseJSON(body.call3) || {
      checked: Boolean(
        body.call3_checked === "true" || body.call3_checked === true
      ),
      notes: body.call3_notes ?? existing.call3?.notes ?? "",
    };

    // normalize category inputs
    const categoryRaw = normalizeField(body.category);
    const categoryOtherRaw = normalizeField(body.categoryOther);

    // compute category to save: if new category provided in request use that; else keep existing
    let updatedCategory = existing.category || "";
    let updatedCategoryOther = existing.categoryOther || "";

    if (categoryRaw) {
      updatedCategory =
        categoryRaw === "other" ? categoryOtherRaw || "" : categoryRaw;
      updatedCategoryOther =
        categoryRaw === "other" ? categoryOtherRaw || "" : "";
    }

    const update = {
      month: normalizeField(body.month) || existing.month,
      quarter: normalizeField(body.quarter) || existing.quarter,
      prospect: normalizeField(body.prospect) || existing.prospect,
      geo: normalizeField(body.geo) || existing.geo,
      lob: normalizeField(body.lob) || existing.lob,

      call1,
      call2,
      call3,

      coreOfferings:
        normalizeField(body.coreOfferings) || existing.coreOfferings,
      primaryNeed: normalizeField(body.primaryNeed) || existing.primaryNeed,
      secondaryNeed:
        normalizeField(body.secondaryNeed) || existing.secondaryNeed,

      category: updatedCategory,
      categoryOther: updatedCategoryOther,

      trace: normalizeField(body.trace) || existing.trace,
      salesSpoc: normalizeField(body.salesSpoc) || existing.salesSpoc,
      oppId: existing.oppId,
      oppDetails: normalizeField(body.oppDetails) || existing.oppDetails,

      deck: deckUrl,
      deckPublicId,
      rag: normalizeField(body.rag) || existing.rag,
      remark: normalizeField(body.remark) || existing.remark,
      updatedAt: new Date(),
    };

    const updated = await ProspectDetail.findByIdAndUpdate(id, update, {
      new: true,
    });
    return res.json({ success: true, data: updated });
  } catch (err) {
    console.error("updateProspect error:", err);
    return res.status(500).json({ success: false, message: "Server error" });
  }
};

exports.downloadProspectDetails = async (req, res, next) => {
  try {
    const { geo, month, quarter, rag } = req.query;

    const query = {};
    if (geo) query.geo = geo;
    if (month) query.month = month;
    if (quarter) query.quarter = quarter;
    if (rag) query.rag = rag;

    const items = await ProspectDetail.find(query).sort({ createdAt: -1 });

    // If no data
    if (!items.length)
      return res
        .status(404)
        .json({ success: false, message: "No records found" });

    //-------------------- EXCEL SETUP --------------------
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Prospects");

    // Header Columns
    const columns = [
      { header: "Prospect", key: "prospect", width: 25 },
      { header: "Geo", key: "geo", width: 15 },
      { header: "Month", key: "month", width: 12 },
      { header: "Quarter", key: "quarter", width: 12 },
      { header: "LOB", key: "lob", width: 20 },
      { header: "Call 1 ", key: "call1Notes", width: 30 },
      { header: "Call 2 ", key: "call2Notes", width: 30 },
      { header: "Call 3 ", key: "call3Notes", width: 30 },
      { header: "Core Offerings", key: "coreOfferings", width: 30 },
      { header: "Primary Need", key: "primaryNeed", width: 30 },
      { header: "Secondary Need", key: "secondaryNeed", width: 30 },
      { header: "Category", key: "category", width: 30 },
      { header: "Trace", key: "trace", width: 30 },
      { header: "Sales SPOC", key: "salesSpoc", width: 30 },
      { header: "Opp ID", key: "oppId", width: 30 },
      { header: "Opp Details", key: "oppDetails", width: 30 },
      { header: "Deck", key: "deck", width: 30 },
      { header: "RAG", key: "rag", width: 10 },
      { header: "Remark", key: "remark", width: 40 },
      { header: "Created At", key: "createdAt", width: 22 },
    ];

    worksheet.columns = columns;

    //-------------------- HEADER STYLE --------------------
    worksheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 12 };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF000000" }, // black
      };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
        bottom: { style: "thin" },
      };
    });

    //-------------------- ADD ROWS --------------------
    // items.forEach((item) => {
    //   worksheet.addRow({
    //     prospect: item.prospect,
    //     geo: item.geo,
    //     month: item.month,
    //     quarter: item.quarter,
    //     lob: item.lob,
    //     category: item.category,
    //     categoryOther: item.categoryOther,
    //     coreOfferings: item.coreOfferings,
    //     primaryNeed: item.primaryNeed,
    //     secondaryNeed: item.secondaryNeed,
    //     call1: item.call1?.checked ? "Yes" : "No",
    //     call1Notes: item.call1?.notes || "",
    //     call2: item.call2?.checked ? "Yes" : "No",
    //     call2Notes: item.call2?.notes || "",
    //     call3: item.call3?.checked ? "Yes" : "No",
    //     call3Notes: item.call3?.notes || "",
    //     trace:item.trace,
    //     salesSpoc:item.salesSpoc,
    //     oppId:item.oppId,
    //     oppDetails:item.oppDetails,
    //     deck:item.deck,
    //     rag: item.rag,
    //     remark: item.remark,
    //     createdAt: item.createdAt?.toISOString()?.slice(0, 10),
    //   });
    // });

    //-------------------- ADD ROWS --------------------
items.forEach((item, index) => {
  const rowIndex = index + 2; // +2 because row 1 = header

  worksheet.addRow({
    prospect: item.prospect,
    geo: item.geo,
    month: item.month,
    quarter: item.quarter,
    lob: item.lob,
    category: item.category,
    categoryOther: item.categoryOther,
    coreOfferings: item.coreOfferings,
    primaryNeed: item.primaryNeed,
    secondaryNeed: item.secondaryNeed,
    call1: item.call1?.checked ? "Yes" : "No",
    call1Notes: item.call1?.notes || "",
    call2: item.call2?.checked ? "Yes" : "No",
    call2Notes: item.call2?.notes || "",
    call3: item.call3?.checked ? "Yes" : "No",
    call3Notes: item.call3?.notes || "",
    trace: item.trace,
    salesSpoc: item.salesSpoc,
    oppId: item.oppId,
    oppDetails: item.oppDetails,
    // deck is handled separately
    rag: item.rag,
    remark: item.remark,
    createdAt: item.createdAt?.toISOString()?.slice(0, 10),
  });

  // Make Deck clickable hyperlink if present
  if (item.deck) {
    const cell = worksheet.getCell(`Q${rowIndex}`); // Q = 17th column (Deck)
    cell.value = {
      text: item.deck,
      hyperlink: item.deck, // URL from DB
    };
    cell.font = { color: { argb: "FF0000FF" }, underline: true }; // blue and underlined
  }
});

    //-------------------- BORDER FOR ALL DATA CELLS --------------------
    worksheet.eachRow((row, rowNum) => {
      if (rowNum === 1) return; // skip header
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
          bottom: { style: "thin" },
        };
      });
    });

    //-------------------- SEND FILE --------------------
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=prospect_data.xlsx`
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    next(err);
  }
};

exports.getProspectDetails = async (req, res, next) => {
  try {
    const {
      page = 1,
      limit = 20,
      search = "",
      geo,
      month,
      quarter,
      // lob,
      rag,
    } = req.query;

    const query = {};

    // Text Search (Prospect name)
    if (search) {
      query.prospect = { $regex: search, $options: "i" };
    }

    // Filters
    if (geo) query.geo = geo;
    if (month) query.month = month;
    if (quarter) query.quarter = quarter;
    // if (lob) query.lob = lob;
    if (rag) query.rag = rag;

    // Pagination
    const pageNum = parseInt(page, 10) || 1;
    const limitNum = parseInt(limit, 10) || 20;
    const skip = (pageNum - 1) * limitNum;

    // Fetch Data
    const [items, total] = await Promise.all([
      ProspectDetail.find(query)
        .sort({ createdAt: -1 })
        .skip(skip)
        .limit(limitNum),

      ProspectDetail.countDocuments(query),
    ]);

    return res.json({
      success: true,
      data: items,
      meta: {
        page: pageNum,
        limit: limitNum,
        total,
      },
    });
  } catch (err) {
    next(err);
  }
};

exports.getCategoryGeoChartData = async (req, res, next) => {
  try {
    const result = await ProspectDetail.aggregate([
      // 1. Only needed fields
      {
        $project: {
          geo: 1,
          category: 1,
        },
      },

      // 2. Remove empty/null fields
      {
        $match: {
          geo: { $exists: true, $ne: "" },
          category: { $exists: true, $ne: "" },
        },
      },

      // 3. Count records by geo + category
      {
        $group: {
          _id: { geo: "$geo", category: "$category" },
          count: { $sum: 1 },
        },
      },

      // 4. Group again by geo, collecting all category-count pairs
      {
        $group: {
          _id: "$_id.geo",
          categoryCounts: {
            $push: {
              category: "$_id.category",
              count: "$count",
            },
          },
        },
      },

      // 5. Final clean object
      {
        $project: {
          geo: "$_id",
          categoryCounts: 1,
          _id: 0,
        },
      },
    ]);

    // Handle empty DB safely
    if (!result || result.length === 0) {
      return res.json({
        success: true,
        xAxis: [],
        seriesLabels: [],
        data: {},
      });
    }

    // Unique geo labels → these will become seriesLabels
    const seriesLabels = result.map((r) => r.geo);

    // Unique categories → these will become xAxis
    const xAxis = [
      ...new Set(
        result.flatMap((r) => (r.categoryCounts || []).map((c) => c.category))
      ),
    ].sort(); // optional: sort for consistency

    // Build data object
    const data = {};
    seriesLabels.forEach((geo) => {
      const geoRecord = result.find((r) => r.geo === geo);
      const categoryCounts = geoRecord?.categoryCounts || [];
      data[geo] = xAxis.map((cat) => {
        const found = categoryCounts.find((c) => c.category === cat);
        return found ? found.count : 0;
      });
    });

    // Send response
    return res.json({
      success: true,
      xAxis,
      seriesLabels,
      data,
      filterNames: ["Categories", "Geo"],
    });
  } catch (err) {
    console.error("Category-Geo Chart Error:", err);
    next(err);
  }
};

// exports.getCategoryMonthChartData = async (req, res, next) => {
//   try {
//     const result = await ProspectDetail.aggregate([
//       // 1. Only required fields
//       {
//         $project: {
//           category: 1,
//           month: 1,
//         },
//       },

//       // 2. Remove empty values
//       {
//         $match: {
//           category: { $exists: true, $ne: "" },
//           month: { $exists: true, $ne: "" },
//         },
//       },

//       // 3. Group by category + month
//       {
//         $group: {
//           _id: { category: "$category", month: "$month" },
//           count: { $sum: 1 },
//         },
//       },

//       // 4. Group again by category → collect all month-count pairs
//       {
//         $group: {
//           _id: "$_id.category",
//           monthCounts: {
//             $push: {
//               month: "$_id.month",
//               count: "$count",
//             },
//           },
//         },
//       },

//       // 5. Final formatting
//       {
//         $project: {
//           category: "$_id",
//           monthCounts: 1,
//           _id: 0,
//         },
//       },
//     ]);

//     // Handle empty DB
//     if (!result || result.length === 0) {
//       return res.json({
//         success: true,
//         xAxis: [],
//         seriesLabels: [],
//         data: {},
//       });
//     }

//     // Unique categories
//     const seriesLabels = result.map((r) => r.category);

//     // Unique months (sorted)
//     const xAxis = [
//       ...new Set(result.flatMap((r) => r.monthCounts.map((m) => m.month))),
//     ].sort((a, b) => {
//       const monthOrder = [
//         "January",
//         "February",
//         "March",
//         "April",
//         "May",
//         "June",
//         "July",
//         "August",
//         "September",
//         "October",
//         "November",
//         "December",
//       ];
//       return monthOrder.indexOf(a) - monthOrder.indexOf(b);
//     });

//     // Build final data object
//     const data = {};
//     seriesLabels.forEach((category) => {
//       const record = result.find((r) => r.category === category);
//       const monthCounts = record?.monthCounts || [];
//       data[category] = xAxis.map((month) => {
//         const found = monthCounts.find((m) => m.month === month);
//         return found ? found.count : 0;
//       });
//     });

//     // Return response with renamed keys
//     return res.json({
//       success: true,
//       xAxis,
//       seriesLabels,
//       data,
//       filterNames: ["Month","Categories"],
//     });
//   } catch (err) {
//     console.error("Category-Month Chart Error:", err);
//     next(err);
//   }
// };

exports.getCategoryMonthChartData = async (req, res, next) => {
  try {
    const result = await ProspectDetail.aggregate([
      // 1. Only required fields
      {
        $project: {
          category: 1,
          month: 1,
        },
      },

      // 2. Remove empty values
      {
        $match: {
          category: { $exists: true, $ne: "" },
          month: { $exists: true, $ne: "" },
        },
      },

      // 3. Group by category + month
      {
        $group: {
          _id: { category: "$category", month: "$month" },
          count: { $sum: 1 },
        },
      },

      // 4. Group again by category → collect all month-count pairs
      {
        $group: {
          _id: "$_id.category",
          monthCounts: {
            $push: {
              month: "$_id.month",
              count: "$count",
            },
          },
        },
      },

      // 5. Final formatting
      {
        $project: {
          category: "$_id",
          monthCounts: 1,
          _id: 0,
        },
      },
    ]);

    // Handle empty DB
    if (!result || result.length === 0) {
      return res.json({
        success: true,
        xAxis: [],
        seriesLabels: [],
        data: {},
      });
    }

    // Fixed month order for sorting
    const monthOrder = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December",
    ];

    // Unique categories
    const seriesLabels = result.map((r) => r.category);

    // Unique, sorted months
    const xAxis = [
      ...new Set(result.flatMap((r) => r.monthCounts.map((m) => m.month))),
    ].sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));

    // Build data object (each category gets aligned month values)
    const data = {};

    seriesLabels.forEach((category) => {
      const record = result.find((r) => r.category === category);
      const monthCounts = record?.monthCounts || [];

      // Convert all categories → sorted month array (fill missing with 0)
      data[category] = xAxis.map((month) => {
        const found = monthCounts.find((m) => m.month === month);
        return found ? found.count : 0;
      });
    });

    // Return fully sorted data
    return res.json({
      success: true,
      xAxis, // sorted months
      seriesLabels, // categories
      data, // aligned month data
      filterNames: ["Month", "Categories"],
    });
  } catch (err) {
    console.error("Category-Month Chart Error:", err);
    next(err);
  }
};

exports.deleteProspect = async (req, res, next) => {
  try {
    const prospect = await ProspectDetail.findById(req.params.id);
    if (!prospect)
      return res.status(404).json({ success: false, message: "Not found" });

    if (prospect.deckPublicId) {
      try {
        await cloudinary.uploader.destroy(prospect.deckPublicId);
      } catch (err) {
        console.error("Cloudinary destroy error:", err);
      }
    }

    await ProspectDetail.findByIdAndDelete(req.params.id);
    res.json({ success: true, message: "Deleted successfully" });
  } catch (err) {
    next(err);
  }
};
