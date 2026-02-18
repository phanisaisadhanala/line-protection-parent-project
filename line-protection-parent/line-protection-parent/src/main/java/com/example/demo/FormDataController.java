package com.example.demo;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.opencsv.CSVReader;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.util.CellRangeAddress;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.*;

@RestController
@CrossOrigin(origins = "http://localhost:8080")
public class FormDataController {

    private static final Logger log = LoggerFactory.getLogger(FormDataController.class);

    @PostMapping(path = "/upload", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<ByteArrayResource> handleUpload(
            @RequestPart("formData") String formDataJson,
            @RequestPart("csvFile") MultipartFile csvFile
    ) throws Exception {

        // Log and parse JSON payload
        log.info("formData (raw json) = {}", formDataJson);
        ObjectMapper mapper = new ObjectMapper();
        Map<String,String> formData = mapper.readValue(formDataJson, new TypeReference<>(){});
        log.info("formData (parsed, {} keys) = {}", formData.size(), formData);
        formData.forEach((k,v)->log.info("Field {} = {}", k, v));

        // Load CSV rows
        List<String[]> csvRows;
        try (CSVReader reader = new CSVReader(
                new java.io.InputStreamReader(csvFile.getInputStream(), StandardCharsets.UTF_8))) {
            csvRows = reader.readAll();
        }
        if (csvRows.isEmpty()) {
            throw new IllegalStateException("CSV empty");
        }
        log.info("CSV rows loaded = {}", csvRows.size());

        // Find section start rows
        int infeedStartRow = findSectionStart(csvRows, "INFEED TAB");
        int impedanceStartRow = findSectionStart(csvRows, "APA IMPEDANCES TAB");

        log.info("Found INFEED TAB at row: {}", infeedStartRow);
        log.info("Found APA IMPEDANCES TAB at row: {}", impedanceStartRow);

        // Parse infeed and impedance data
        Map<Integer, InfeedData> infeedMap = parseInfeedData(csvRows, infeedStartRow, impedanceStartRow);
        ImpedanceData impedanceData = parseImpedanceData(csvRows, impedanceStartRow);

        log.info("Parsed {} infeed entries from CSV", infeedMap.size());
        log.info("Parsed impedance data: First line + {} second lines", impedanceData.secondLines.size());

        // Open template and workbook safely (try-with-resources)
        ClassPathResource tpl = new ClassPathResource("Line Protection Calculation Sheet Template.xlsm");
        if (!tpl.exists()) {
            throw new java.io.FileNotFoundException("Template missing");
        }

        try (InputStream is = tpl.getInputStream();
             OPCPackage pkg = OPCPackage.open(is)){

            sanitizeVml(pkg);


            try(XSSFWorkbook wb = new XSSFWorkbook(pkg);
                ByteArrayOutputStream out = new ByteArrayOutputStream()) {

                // ===== ASSIGNING VALUES IN TAB DATA ENTRY ===== //
                Sheet DataEntryTab = wb.getSheet("1) Data Entry");
                if (DataEntryTab == null) {
                    log.warn("Sheet '1) Data Entry' not found in template");
                } else {
                    log.info("========== DATA ENTRY TAB MAPPING ==========");

                    writeCellMerged(DataEntryTab, "G3", formData.getOrDefault("relayLocation", ""));
                    writeCellMerged(DataEntryTab, "K3", formData.getOrDefault("lineNumber", ""));
                    writeCellMerged(DataEntryTab, "M3", formData.getOrDefault("remoteLocation", ""));
                    writeCellMerged(DataEntryTab, "E16", formData.getOrDefault("nominalSystemVoltage", ""));
                    writeCellMerged(DataEntryTab, "E18", formData.getOrDefault("breakerRating", ""));
                    writeCellMerged(DataEntryTab, "E19", formData.getOrDefault("conductorRating", ""));
                    writeCellMerged(DataEntryTab, "E22", formData.getOrDefault("ctrW", ""));
                    writeCellMerged(DataEntryTab, "E23", formData.getOrDefault("ctrX", ""));
                    writeCellMerged(DataEntryTab, "E24", formData.getOrDefault("ptry", ""));
                    writeCellMerged(DataEntryTab, "E26", formData.getOrDefault("secondlines", ""));
                    writeCellMerged(DataEntryTab, "E27", formData.getOrDefault("numberOfTaps", ""));
                    writeCellMerged(DataEntryTab, "E28", formData.getOrDefault("autoXfmrAtRemote", ""));
                    writeCellMerged(DataEntryTab, "E29", formData.getOrDefault("numberOfBreakers", ""));
                    writeCellMerged(DataEntryTab, "E30", formData.getOrDefault("noOfDistributionTransformers", ""));
                    writeCellMerged(DataEntryTab, "E34", formData.getOrDefault("relayLoadbility", ""));
                    writeCellMerged(DataEntryTab, "E270", formData.getOrDefault("syncReference", ""));
                    writeCellMerged(DataEntryTab, "E271", formData.getOrDefault("syncSource", ""));
                    writeCellMerged(DataEntryTab, "E273", formData.getOrDefault("hotLineInd", ""));
                    writeCellMerged(DataEntryTab, "E274", formData.getOrDefault("vazPtRatio", ""));
                    writeCellMerged(DataEntryTab, "E275", formData.getOrDefault("vbzPtRatio", ""));
                    writeCellMerged(DataEntryTab, "E276", formData.getOrDefault("vczPtRatio", ""));
                    writeCellMerged(DataEntryTab, "E282", formData.getOrDefault("remoteCTR", ""));
                    writeCellMerged(DataEntryTab, "E285", formData.getOrDefault("remoteBFPU", ""));
                    writeCellMerged(DataEntryTab, "E286", formData.getOrDefault("remoteBFGU", ""));

                    log.info("Data Entry tab mapping complete");
                }

                // ===== ASSIGNING VALUES IN TAB FAULT ANALYSIS ===== //
                Sheet FaultAnalysisTab = wb.getSheet("4) Fault Analysis");
                if (FaultAnalysisTab == null) {
                    throw new IllegalStateException("Sheet '4) Fault Analysis' not found in template");
                }

                log.info("========== FAULT ANALYSIS TAB MAPPING ==========");

                // Fault Analysis mappings
                String G17 = csvHandler(csvRows, 0, 2);
                log.info("WRITE G17 (Min Line End SLG All Sources) <= CSV[0][2] : {}", G17.isBlank()?"<EMPTY>":G17);
                writeIfPresent(FaultAnalysisTab, "G17", G17);

                String G18 = csvHandler(csvRows, 1, 2);
                log.info("WRITE G18 (Min Line End 1LG All Sources) <= CSV[1][2] : {}", G18.isBlank()?"<EMPTY>":G18);
                writeIfPresent(FaultAnalysisTab, "G18", G18);

                String G19 = csvHandler(csvRows, 2, 2);
                log.info("WRITE G19 (Min Line End LL I2) <= CSV[2][2] : {}", G19.isBlank()?"<EMPTY>":G19);
                writeIfPresent(FaultAnalysisTab, "G19", G19);

                String G22 = csvHandler(csvRows, 6, 2);
                log.info("WRITE G22 (Min Line End n-1 SLG) <= CSV[6][2] : {}", G22.isBlank()?"<EMPTY>":G22);
                writeIfPresent(FaultAnalysisTab, "G22", G22);

                String G23 = csvHandler(csvRows, 7, 2);
                log.info("WRITE G23 (Min Line End n-1 I2) <= CSV[7][2] : {}", G23.isBlank()?"<EMPTY>":G23);
                writeIfPresent(FaultAnalysisTab, "G23", G23);

                String G25 = csvHandler(csvRows, 8, 2);
                log.info("WRITE G25 (Reverse Local Bus 1LG) <= CSV[8][2] : {}", G25.isBlank()?"<EMPTY>":G25);
                writeIfPresent(FaultAnalysisTab, "G25", G25);

                String CEO3LG = csvHandler(csvRows, 9, 2);
                log.info("WRITE CEO3LG (Close In End Open 3LG) <= CSV[9][2] : {}", CEO3LG.isBlank()?"<EMPTY>":CEO3LG);
                writeIfPresent(FaultAnalysisTab, "G36", CEO3LG);
                String CEO1LG = csvHandler(csvRows, 9, 4);
                log.info("WRITE CEO1LG (Close In End Open 1LG) <= CSV[9][4] : {}", CEO1LG.isBlank()?"<EMPTY>":CEO1LG);
                writeIfPresent(FaultAnalysisTab, "K36", CEO1LG);

                String CEC3LG = csvHandler(csvRows, 10, 2);
                log.info("WRITE CEC3LG (Close In End Closed 3LG) <= CSV[10][2] : {}", CEC3LG.isBlank()?"<EMPTY>":CEC3LG);
                writeIfPresent(FaultAnalysisTab, "G37", CEC3LG);
                String CEC1LG = csvHandler(csvRows, 10, 4);
                log.info("WRITE CEC1LG (Close In End Closed 1LG) <= CSV[10][4] : {}", CEC1LG.isBlank()?"<EMPTY>":CEC1LG);
                writeIfPresent(FaultAnalysisTab, "K37", CEC1LG);

                String SSR3LG = csvHandler(csvRows, 12, 2);
                log.info("WRITE SSR3LG (Remote Bus Fault 3LG) <= CSV[12][2] : {}", SSR3LG.isBlank()?"<EMPTY>":SSR3LG);
                writeIfPresent(FaultAnalysisTab, "G38", SSR3LG);
                String SSRLL = csvHandler(csvRows, 12, 4);
                log.info("WRITE SSRLL (Remote Bus Fault L-L) <= CSV[12][4] : {}", SSRLL.isBlank()?"<EMPTY>":SSRLL);
                writeIfPresent(FaultAnalysisTab, "I38", SSRLL);
                String SSR3I0 = csvHandler(csvRows, 12, 6);
                log.info("WRITE SSR3I0 (Remote Bus Fault 1LG 3IO) <= CSV[12][6] : {}", SSR3I0.isBlank()?"<EMPTY>":SSR3I0);
                writeIfPresent(FaultAnalysisTab, "K38", SSR3I0);

                String L2NDL = csvHandler(csvRows, 13, 2);
                log.info("WRITE L2NDL (Longest 2nd Line SLG 3IO) <= CSV[13][2] : {}", L2NDL.isBlank()?"<EMPTY>":L2NDL);
                writeIfPresent(FaultAnalysisTab, "K39", L2NDL);

                // Differential Relay Calculations
                String DIFF1A3LG = csvHandler(csvRows, 15, 2);
                log.info("WRITE DIFF1A3LG (Diff Case 1a 3LG) <= CSV[15][2] : {}", DIFF1A3LG.isBlank()?"<EMPTY>":DIFF1A3LG);
                writeIfPresent(FaultAnalysisTab, "S35", DIFF1A3LG);
                String DIFF1ALL = csvHandler(csvRows, 15, 4);
                log.info("WRITE DIFF1ALL (Diff Case 1a L-L) <= CSV[15][4] : {}", DIFF1ALL.isBlank()?"<EMPTY>":DIFF1ALL);
                writeIfPresent(FaultAnalysisTab, "U35", DIFF1ALL);
                String DIFF1AI2 = csvHandler(csvRows, 15, 6);
                log.info("WRITE DIFF1AI2 (Diff Case 1a I2) <= CSV[15][6] : {}", DIFF1AI2.isBlank()?"<EMPTY>":DIFF1AI2);
                writeIfPresent(FaultAnalysisTab, "W35", DIFF1AI2);
                String DIFF1A3I0 = csvHandler(csvRows, 15, 8);
                log.info("WRITE DIFF1A3I0 (Diff Case 1a 3I0) <= CSV[15][8] : {}", DIFF1A3I0.isBlank()?"<EMPTY>":DIFF1A3I0);
                writeIfPresent(FaultAnalysisTab, "X35", DIFF1A3I0);

                String DIFF1B3LG = csvHandler(csvRows, 17, 2);
                log.info("WRITE DIFF1B3LG (Diff Case 1b 3LG) <= CSV[17][2] : {}", DIFF1B3LG.isBlank()?"<EMPTY>":DIFF1B3LG);
                writeIfPresent(FaultAnalysisTab, "S36", DIFF1B3LG);
                String DIFF1BLL = csvHandler(csvRows, 17, 4);
                log.info("WRITE DIFF1BLL (Diff Case 1b L-L) <= CSV[17][4] : {}", DIFF1BLL.isBlank()?"<EMPTY>":DIFF1BLL);
                writeIfPresent(FaultAnalysisTab, "U36", DIFF1BLL);
                String DIFF1BI2 = csvHandler(csvRows, 17, 6);
                log.info("WRITE DIFF1BI2 (Diff Case 1b I2) <= CSV[17][6] : {}", DIFF1BI2.isBlank()?"<EMPTY>":DIFF1BI2);
                writeIfPresent(FaultAnalysisTab, "W36", DIFF1BI2);
                String DIFF1B3I0 = csvHandler(csvRows, 17, 8);
                log.info("WRITE DIFF1B3I0 (Diff Case 1b 3I0) <= CSV[17][8] : {}", DIFF1B3I0.isBlank()?"<EMPTY>":DIFF1B3I0);
                writeIfPresent(FaultAnalysisTab, "X36", DIFF1B3I0);

                String DIFF2A3LG = csvHandler(csvRows, 18, 2);
                log.info("WRITE DIFF2A3LG (Diff Case 2a 3LG) <= CSV[18][2] : {}", DIFF2A3LG.isBlank()?"<EMPTY>":DIFF2A3LG);
                writeIfPresent(FaultAnalysisTab, "S37", DIFF2A3LG);
                String DIFF2ALL = csvHandler(csvRows, 18, 4);
                log.info("WRITE DIFF2ALL (Diff Case 2a L-L) <= CSV[18][4] : {}", DIFF2ALL.isBlank()?"<EMPTY>":DIFF2ALL);
                writeIfPresent(FaultAnalysisTab, "U37", DIFF2ALL);
                String DIFF2AI2 = csvHandler(csvRows, 18, 6);
                log.info("WRITE DIFF2AI2 (Diff Case 2a I2) <= CSV[18][6] : {}", DIFF2AI2.isBlank()?"<EMPTY>":DIFF2AI2);
                writeIfPresent(FaultAnalysisTab, "W37", DIFF2AI2);
                String DIFF2A3I0 = csvHandler(csvRows, 18, 8);
                log.info("WRITE DIFF2A3I0 (Diff Case 2a 3I0) <= CSV[18][8] : {}", DIFF2A3I0.isBlank()?"<EMPTY>":DIFF2A3I0);
                writeIfPresent(FaultAnalysisTab, "X37", DIFF2A3I0);

                String DIFF2B3LG = csvHandler(csvRows, 19, 2);
                log.info("WRITE DIFF2B3LG (Diff Case 2b 3LG) <= CSV[19][2] : {}", DIFF2B3LG.isBlank()?"<EMPTY>":DIFF2B3LG);
                writeIfPresent(FaultAnalysisTab, "S38", DIFF2B3LG);
                String DIFF2BLL = csvHandler(csvRows, 19, 4);
                log.info("WRITE DIFF2BLL (Diff Case 2b L-L) <= CSV[19][4] : {}", DIFF2BLL.isBlank()?"<EMPTY>":DIFF2BLL);
                writeIfPresent(FaultAnalysisTab, "U38", DIFF2BLL);
                String DIFF2BI2 = csvHandler(csvRows, 19, 6);
                log.info("WRITE DIFF2BI2 (Diff Case 2b I2) <= CSV[19][6] : {}", DIFF2BI2.isBlank()?"<EMPTY>":DIFF2BI2);
                writeIfPresent(FaultAnalysisTab, "W38", DIFF2BI2);
                String DIFF2B3I0 = csvHandler(csvRows, 19, 8);
                log.info("WRITE DIFF2B3I0 (Diff Case 2b 3I0) <= CSV[19][8] : {}", DIFF2B3I0.isBlank()?"<EMPTY>":DIFF2B3I0);
                writeIfPresent(FaultAnalysisTab, "X38", DIFF2B3I0);

                // X/R n-0
                String XRNminus03LG = csvHandler(csvRows, 21, 2);
                log.info("WRITE XRNminus03LG (X/R N-0 3LG) <= CSV[21][2] : {}", XRNminus03LG.isBlank()?"<EMPTY>":XRNminus03LG);
                writeIfPresent(FaultAnalysisTab, "E44", XRNminus03LG);

                String XRNminus01LG = csvHandler(csvRows, 22, 2);
                log.info("WRITE XRNminus01LG (X/R N-0 SLG) <= CSV[22][2] : {}", XRNminus01LG.isBlank()?"<EMPTY>":XRNminus01LG);
                writeIfPresent(FaultAnalysisTab, "E45", XRNminus01LG);

                String XRNminus0R1 = csvHandler(csvRows, 23, 2);
                log.info("WRITE XRNminus0R1 (X/R N-0 R1) <= CSV[23][2] : {}", XRNminus0R1.isBlank()?"<EMPTY>":XRNminus0R1);
                writeIfPresent(FaultAnalysisTab, "E47", XRNminus0R1);
                String XRNminus0X1 = csvHandler(csvRows, 23, 4);
                log.info("WRITE XRNminus0X1 (X/R N-0 X1) <= CSV[23][4] : {}", XRNminus0X1.isBlank()?"<EMPTY>":XRNminus0X1);
                writeIfPresent(FaultAnalysisTab, "G47", XRNminus0X1);
                String XRNminus0R2 = csvHandler(csvRows, 23, 6);
                log.info("WRITE XRNminus0R2 (X/R N-0 R2) <= CSV[23][6] : {}", XRNminus0R2.isBlank()?"<EMPTY>":XRNminus0R2);
                writeIfPresent(FaultAnalysisTab, "I47", XRNminus0R2);
                String XRNminus0X2 = csvHandler(csvRows, 23, 8);
                log.info("WRITE XRNminus0X2 (X/R N-0 X2) <= CSV[23][8] : {}", XRNminus0X2.isBlank()?"<EMPTY>":XRNminus0X2);
                writeIfPresent(FaultAnalysisTab, "K47", XRNminus0X2);
                String XRNminus0R0 = csvHandler(csvRows, 23, 10);
                log.info("WRITE XRNminus0R0 (X/R N-0 R0) <= CSV[23][10] : {}", XRNminus0R0.isBlank()?"<EMPTY>":XRNminus0R0);
                writeIfPresent(FaultAnalysisTab, "M47", XRNminus0R0);
                String XRNminus0X0 = csvHandler(csvRows, 23, 12);
                log.info("WRITE XRNminus0X0 (X/R N-0 X0) <= CSV[23][12] : {}", XRNminus0X0.isBlank()?"<EMPTY>":XRNminus0X0);
                writeIfPresent(FaultAnalysisTab, "O47", XRNminus0X0);

                // X/R n-1
                String XRNminus13LG = csvHandler(csvRows, 24, 2);
                log.info("WRITE XRNminus13LG (X/R N-1 3LG) <= CSV[24][2] : {}", XRNminus13LG.isBlank()?"<EMPTY>":XRNminus13LG);
                writeIfPresent(FaultAnalysisTab, "E70", XRNminus13LG);

                String XRNminus11LG = csvHandler(csvRows, 25, 2);
                log.info("WRITE XRNminus11LG (X/R N-1 SLG) <= CSV[25][2] : {}", XRNminus11LG.isBlank()?"<EMPTY>":XRNminus11LG);
                writeIfPresent(FaultAnalysisTab, "E71", XRNminus11LG);

                String XRNminus1R1 = csvHandler(csvRows, 26, 2);
                log.info("WRITE XRNminus1R1 (X/R N-1 R1) <= CSV[26][2] : {}", XRNminus1R1.isBlank()?"<EMPTY>":XRNminus1R1);
                writeIfPresent(FaultAnalysisTab, "E73", XRNminus1R1);
                String XRNminus1X1 = csvHandler(csvRows, 26, 4);
                log.info("WRITE XRNminus1X1 (X/R N-1 X1) <= CSV[26][4] : {}", XRNminus1X1.isBlank()?"<EMPTY>":XRNminus1X1);
                writeIfPresent(FaultAnalysisTab, "G73", XRNminus1X1);
                String XRNminus1R2 = csvHandler(csvRows, 26, 6);
                log.info("WRITE XRNminus1R2 (X/R N-1 R2) <= CSV[26][6] : {}", XRNminus1R2.isBlank()?"<EMPTY>":XRNminus1R2);
                writeIfPresent(FaultAnalysisTab, "I73", XRNminus1R2);
                String XRNminus1X2 = csvHandler(csvRows, 26, 8);
                log.info("WRITE XRNminus1X2 (X/R N-1 X2) <= CSV[26][8] : {}", XRNminus1X2.isBlank()?"<EMPTY>":XRNminus1X2);
                writeIfPresent(FaultAnalysisTab, "K73", XRNminus1X2);
                String XRNminus1R0 = csvHandler(csvRows, 26, 10);
                log.info("WRITE XRNminus1R0 (X/R N-1 R0) <= CSV[26][10] : {}", XRNminus1R0.isBlank()?"<EMPTY>":XRNminus1R0);
                writeIfPresent(FaultAnalysisTab, "M73", XRNminus1R0);
                String XRNminus1X0 = csvHandler(csvRows, 26, 12);
                log.info("WRITE XRNminus1X0 (X/R N-1 X0) <= CSV[26][12] : {}", XRNminus1X0.isBlank()?"<EMPTY>":XRNminus1X0);
                writeIfPresent(FaultAnalysisTab, "O73", XRNminus1X0);

                log.info("Fault Analysis tab mapping complete");

                // ===== ASSIGNING VALUES IN TAB INFEED ===== //
                Sheet InfeedTab = wb.getSheet("5) Infeed");
                if (InfeedTab == null) {
                    log.warn("Sheet '5) Infeed' not found in template");
                } else {
                    log.info("========== INFEED TAB MAPPING (Always fill all 12 buses, default to 0) ==========");

                    for (int busNum = 1; busNum <= 12; busNum++) {
                        int excelRow = 14 + busNum;
                        String magCell = "R" + excelRow;
                        String angCell = "T" + excelRow;

                        InfeedData data = infeedMap.getOrDefault(busNum, new InfeedData("0", "0"));

                        log.info("Bus {}: Mag={} -> {}, Ang={} -> {}", busNum, data.magnitude, magCell, data.angle, angCell);

                        writeCellMerged(InfeedTab, magCell, data.magnitude);
                        writeCellMerged(InfeedTab, angCell, data.angle);
                    }

                    log.info("Infeed tab mapping complete (all 12 buses filled)");
                }

                // ===== ASSIGNING VALUES IN TAB ASPEN IMPEDANCES ===== //
                Sheet APAImpedancesTab = wb.getSheet("3) Aspen Impedances");
                if (APAImpedancesTab == null) {
                    log.warn("Sheet '3) Aspen Impedances' not found in template");
                } else {
                    log.info("========== ASPEN IMPEDANCES TAB MAPPING ==========");

                    // First Line Impedance
                    if (impedanceData.firstLine != null) {
                        log.info("WRITE FirstLineImpedance E6={}, F6={}, G6={}, H6={}, I6={}",
                                impedanceData.firstLine.r1, impedanceData.firstLine.x1,
                                impedanceData.firstLine.r0, impedanceData.firstLine.x0,
                                impedanceData.firstLine.miles);

                        writeIfPresent(APAImpedancesTab, "E6", impedanceData.firstLine.r1);
                        writeIfPresent(APAImpedancesTab, "F6", impedanceData.firstLine.x1);
                        writeIfPresent(APAImpedancesTab, "G6", impedanceData.firstLine.r0);
                        writeIfPresent(APAImpedancesTab, "H6", impedanceData.firstLine.x0);
                        writeIfPresent(APAImpedancesTab, "I6", impedanceData.firstLine.miles);
                    }

                    // -----------------------------------------------------------------------
                    // Second Line Impedances
                    //
                    // Template layout per section (8 sections total):
                    //
                    //   excelRow        → primary data entry row      (E/F/G/H/I columns)
                    //   excelRow + 1..4 → green formula/display rows  (auto-calculated)
                    //   excelRow + 5    → "Not Used" summary row      ← MUST also be written
                    //   excelRow + 6    → "Bus CAPE CKT number" row   ← D column gets CKT ID
                    //   excelRow + 7    → section separator / empty
                    //
                    // Without explicitly writing the "Not Used" row (excelRow+5), the
                    // template retains its previous / default value, causing the mismatch
                    // visible in the screenshot (e.g. row 44 showing YANDELL-26 data
                    // instead of HOYRD data).
                    // -----------------------------------------------------------------------
                    int[] excelRows = {15, 23, 31, 39, 47, 55, 63, 71}; // primary data rows for 8 second lines

                    for (int i = 0; i < Math.min(impedanceData.secondLines.size(), excelRows.length); i++) {
                        LineImpedance line = impedanceData.secondLines.get(i);
                        int excelRow    = excelRows[i];
                        int yellowCell  = excelRow + 5; // yellow blank cell → write CKT name e.g. "YANDELL-22"

                        log.info("Second Line {} => primaryRow={}, yellowCellRow={} | " +
                                        "R1={} X1={} R0={} X0={} Miles={} CKT='{}'",
                                i + 1, excelRow, yellowCell,
                                line.r1, line.x1, line.r0, line.x0, line.miles, line.cktNumber);

                        // Write impedance values to the primary data-entry row only
                        writeIfPresent(APAImpedancesTab, "E" + excelRow, line.r1);
                        writeIfPresent(APAImpedancesTab, "F" + excelRow, line.x1);
                        writeIfPresent(APAImpedancesTab, "G" + excelRow, line.r0);
                        writeIfPresent(APAImpedancesTab, "H" + excelRow, line.x0);
                        writeIfPresent(APAImpedancesTab, "I" + excelRow, line.miles);

                        // Write CKT name (e.g. "YANDELL-22") into the yellow cell at D(excelRow+5)
                        if (line.cktNumber != null && !line.cktNumber.isBlank()) {
                            log.info("  Writing CKT name '{}' to D{}", line.cktNumber, yellowCell);
                            writeCellMerged(APAImpedancesTab, "D" + yellowCell, line.cktNumber);
                        }
                    }

                    log.info("Aspen Impedances tab mapping complete");
                }

                // Ask Excel to do a full recalc when the user opens the file
                wb.setForceFormulaRecalculation(true);

                log.info("========== ALL MAPPING COMPLETE - Writing workbook ==========");
                wb.write(out);

                ByteArrayResource resource = new ByteArrayResource(out.toByteArray());
                return ResponseEntity.ok()
                        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Updated_Line_Protection_Calculation_Sheet.xlsm")
                        .contentType(MediaType.parseMediaType("application/vnd.ms-excel.sheet.macroEnabled.12"))
                        .body(resource);
            }
        }
    }

    /**
     * Find the starting row of a section in the CSV
     */
    private static int findSectionStart(List<String[]> csvRows, String sectionName) {
        for (int i = 0; i < csvRows.size(); i++) {
            String line = String.join(",", csvRows.get(i)).trim().toUpperCase();
            if (line.contains(sectionName.toUpperCase())) {
                return i;
            }
        }
        return -1;
    }

    /**
     * Parse infeed data from CSV file starting from infeedStartRow
     */
    private static Map<Integer, InfeedData> parseInfeedData(List<String[]> csvRows, int infeedStartRow, int impedanceStartRow) {
        Map<Integer, InfeedData> infeedMap = new HashMap<>();

        if (infeedStartRow == -1) {
            log.warn("INFEED TAB section not found in CSV");
            return infeedMap;
        }

        int busCounter = 1;
        int endRow = (impedanceStartRow != -1) ? impedanceStartRow : csvRows.size();

        for (int i = infeedStartRow; i < endRow && busCounter <= 12; i++) {
            String line = String.join(",", csvRows.get(i)).trim();

            if (line.toUpperCase().contains("WHEN APPLYING BUS FAULT AT:")) {
                String magnitude = "0";
                String angle = "0";

                if (i + 1 < endRow) {
                    String magLine = csvHandler(csvRows, i + 1, 2);
                    if (!magLine.isBlank()) magnitude = magLine;
                }

                if (i + 2 < endRow) {
                    String angLine = csvHandler(csvRows, i + 2, 2);
                    if (!angLine.isBlank()) angle = angLine;
                }

                infeedMap.put(busCounter, new InfeedData(magnitude, angle));
                log.debug("Parsed infeed bus {}: Mag={}, Ang={}", busCounter, magnitude, angle);
                busCounter++;
                i += 2;
            }
        }

        return infeedMap;
    }

    /**
     * Parse impedance data from CSV file starting from impedanceStartRow
     */
    private static ImpedanceData parseImpedanceData(List<String[]> csvRows, int impedanceStartRow) {
        ImpedanceData impedanceData = new ImpedanceData();

        if (impedanceStartRow == -1) {
            log.warn("APA IMPEDANCES TAB section not found in CSV");
            return impedanceData;
        }

        for (int i = impedanceStartRow; i < csvRows.size(); i++) {
            String line = String.join(",", csvRows.get(i)).trim();

            if (line.toUpperCase().contains("FIRST LINE IMPEDENCE AT BUS:")) {
                LineImpedance firstLine = new LineImpedance();

                if (i + 1 < csvRows.size()) {
                    firstLine.r1 = firstNumber(csvHandler(csvRows, i + 1, 2));
                    firstLine.x1 = firstNumber(csvHandler(csvRows, i + 1, 3));
                }
                if (i + 2 < csvRows.size()) {
                    firstLine.r0 = firstNumber(csvHandler(csvRows, i + 2, 2));
                    firstLine.x0 = firstNumber(csvHandler(csvRows, i + 2, 3));
                }
                if (i + 3 < csvRows.size()) {
                    firstLine.miles = firstNumber(csvHandler(csvRows, i + 3, 2));
                }

                impedanceData.firstLine = firstLine;
                log.debug("Parsed first line impedance: R1={}, X1={}, R0={}, X0={}, Miles={}",
                        firstLine.r1, firstLine.x1, firstLine.r0, firstLine.x0, firstLine.miles);
                i += 3;
            }
            else if (line.toUpperCase().contains("SECOND LINE IMPEDENCES FOR LINE:")) {
                LineImpedance secondLine = new LineImpedance();

                String[] lineFields = csvRows.get(i);
                for (String field : lineFields) {
                    if (field != null && field.toUpperCase().contains("SECOND LINE IMPEDENCES FOR LINE:")) {
                        String label = field.trim();
                        int colonIdx = label.lastIndexOf(':');
                        if (colonIdx >= 0 && colonIdx + 1 < label.length()) {
                            String fullId = label.substring(colonIdx + 1).trim(); // e.g. "5586-YANDELL-22"
                            int dashIdx = fullId.indexOf('-');
                            // cktNumber = everything after the first dash e.g. "YANDELL-22"
                            secondLine.cktNumber = (dashIdx >= 0 && dashIdx + 1 < fullId.length())
                                    ? fullId.substring(dashIdx + 1).trim()
                                    : fullId;
                        }
                        break;
                    }
                }
                log.debug("Parsed CAPE CKT number for second line {}: {}", impedanceData.secondLines.size() + 1, secondLine.cktNumber);

                if (i + 1 < csvRows.size()) {
                    secondLine.r1 = firstNumber(csvHandler(csvRows, i + 1, 2));
                    secondLine.x1 = firstNumber(csvHandler(csvRows, i + 1, 3));
                }
                if (i + 2 < csvRows.size()) {
                    secondLine.r0 = firstNumber(csvHandler(csvRows, i + 2, 2));
                    secondLine.x0 = firstNumber(csvHandler(csvRows, i + 2, 3));
                }
                if (i + 3 < csvRows.size()) {
                    secondLine.miles = firstNumber(csvHandler(csvRows, i + 3, 2));
                }

                impedanceData.secondLines.add(secondLine);
                log.debug("Parsed second line {} impedance: R1={}, X1={}, R0={}, X0={}, Miles={}",
                        impedanceData.secondLines.size(), secondLine.r1, secondLine.x1,
                        secondLine.r0, secondLine.x0, secondLine.miles);
                i += 3;
            }
        }

        return impedanceData;
    }

    private static class InfeedData {
        String magnitude;
        String angle;

        InfeedData(String magnitude, String angle) {
            this.magnitude = magnitude;
            this.angle = angle;
        }
    }

    private static class LineImpedance {
        String r1 = "";
        String x1 = "";
        String r0 = "";
        String x0 = "";
        String miles = "";
        String cktNumber = ""; // CKT name written to yellow cell (col D, excelRow+5), e.g. "YANDELL-22"
    }

    private static class ImpedanceData {
        LineImpedance firstLine;
        List<LineImpedance> secondLines = new ArrayList<>();
    }

    private static void sanitizeVml(OPCPackage pkg) throws Exception {
        final String VML_CT = "application/vnd.openxmlformats-officedocument.vmlDrawing";
        for (PackagePart part : pkg.getPartsByContentType(VML_CT)) {
            String xml;
            try (InputStream in = part.getInputStream()) {
                xml = new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
            }

            String fixed = xml;
            fixed = fixed.replaceAll("(?i)<font\\s*>", "<font/>");
            fixed = fixed.replaceAll("&(?![#a-zA-Z0-9]+;)", "&amp;");

            if (!fixed.equals(xml)) {
                try (OutputStream out = part.getOutputStream()) {
                    out.write(fixed.getBytes(java.nio.charset.StandardCharsets.UTF_8));
                }
            }
        }
    }

    private static void writeCellMerged(Sheet sheet, String addr, String raw) {
        CellAddress ca = new CellAddress(addr);
        int r = ca.getRow();
        int c = ca.getColumn();

        CellRangeAddress merged = null;
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress rng = sheet.getMergedRegion(i);
            if (rng.isInRange(r, c)) {
                merged = rng;
                break;
            }
        }
        int wr = (merged != null) ? merged.getFirstRow()    : r;
        int wc = (merged != null) ? merged.getFirstColumn() : c;

        Row row = sheet.getRow(wr);
        if (row == null) row = sheet.createRow(wr);
        Cell cell = row.getCell(wc, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

        try {
            cell.setCellValue(Double.parseDouble(raw));
        } catch (Exception e) {
            cell.setCellValue(raw == null ? "" : raw);
        }
    }

    private static String csvHandler(List<String[]> rows, int r, int c) {
        if (rows == null || r < 0 || r >= rows.size()) return "";
        String[] row = rows.get(r);
        if (row == null || c < 0 || c >= row.length) return "";
        String v = row[c];
        if (v == null) return "";
        return v.replace('\u00A0',' ').trim();
    }

    private static void writeIfPresent(Sheet sheet, String addr, String val) {
        if (val != null && !val.isBlank()) {
            writeCellMerged(sheet, addr, val);
        }
    }

    private static String firstNumber(String s) {
        if (s == null) return "";

        String t = s
                .replace('\u00A0', ' ')
                .replace(",", "")
                .trim();

        String upper = t.toUpperCase();

        if (upper.contains("INFINITE")
                || upper.contains("INFINITY")
                || upper.matches(".*\\bINF\\b.*")
                || t.contains("∞")) {
            return "0";
        }

        java.util.regex.Matcher m = java.util.regex.Pattern
                .compile("[-+]?\\d*\\.?\\d+(?:[eE][-+]?\\d+)?")
                .matcher(t);

        return m.find() ? m.group() : "";
    }
}