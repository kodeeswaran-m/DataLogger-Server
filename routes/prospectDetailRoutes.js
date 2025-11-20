// routes/prospectDetailRoutes.js
const express = require("express");
const router = express.Router();
const controller = require("../controllers/prospectDetailController");
const upload = require("../middleware/upload");


router.get("/chart-data", controller.getCategoryGeoChartData);

router.get("/chart-data-month", controller.getCategoryMonthChartData);

// POST /api/prospectDetail → with file upload (field name: deck)
router.post("/", upload.single("deck"), controller.createProspectDetail);

// GET /api/prospectDetail → list (supports ?page=&limit=&prospect=&geo=&rag=)
router.get("/", controller.getProspectDetails);

// GET /api/prospectDetail/:id → get single
router.get("/:id", controller.getProspectById);

// PUT /api/prospectDetail/:id → update (accept file as multipart/form-data)
router.put("/:id", upload.single("deck"), controller.updateProspect);

// DELETE /api/prospectDetail/:id → delete
router.delete("/:id", controller.deleteProspect);

module.exports = router;
