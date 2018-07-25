/*=============================================================================
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *                                                                             *
 *                    COPYRIGHT, 2017 FORD MOTOR COMPANY                       *
 *                                                                             *
 *                               CONFIDENTIAL                                  *
 *                                                                             *
 * This is an unpublished work, which is a trade secret, created in            *
 * 2016.  Ford Motor Company owns all rights to this work and intends          *
 * to maintain it in confidence to preserve its trade secret status.           *
 * Ford Motor Company reserves the right to protect this work as an            *
 * unpublished copyrighted work in the event of an inadvertent or              *
 * deliberate unauthorized publication. Ford Motor Company also                *
 * reserves its rights under the copyright laws to protect this work           *
 * as a published work. Those having access to this work may not copy          *
 * it, use it, or disclose the information contained in it without the         *
 * written authorization of Ford Motor Company.                                *
 *                                                                             *
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
  File Name: FillInTemplate.java
  File description:
 ==============================================================================
 $HISTORY$
 ------------------------------------------------------------------------------
/*=======================================================================
Date            Name                    Description of Change
June-8-2017     HDAOUD                	US296418 - Print to Excel from CMF - F26344
Oct-16-2017     VTIRUPAT                Find Bugs & US373638-UI/UX -Export to Excel file naming convention
=======================================================================*/

package com.ford.pd.bom.ui.utils;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TimeZone;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.regex.Pattern;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.io.FileExistsException;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontFamily;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;

import com.ford.pd.bom.business.service.ImpactAssessmentConstants;
import com.ford.pd.bom.business.util.BomUtil;
import com.ford.pd.bom.business.util.CmfWorkFlowUtil;
import com.ford.pd.bom.common.cmf.model.CRDetailModel;
import com.ford.pd.bom.common.cmf.util.CmfConstants;
import com.ford.pd.bom.common.tc.FB4ChangeGroup;
import com.ford.pd.bom.common.tc.FB4ChangeRequest;
import com.ford.pd.bom.common.tc.FB4GS2;
import com.ford.pd.bom.common.tc.FB4ManagedEvent;
import com.ford.pd.bom.common.util.BCMExcelConstants;
import com.ford.pd.bom.common.util.BomAuthoringConstants;
import com.ford.pd.bom.common.util.BomSearchConstants;
import com.ford.pd.bom.common.util.Tuple;
import com.ford.pd.bom.common.vo.ImpactAssesmentVO;
import com.ford.pd.bom.domain.data.manager.FB4BOMPrgIndependentDataHolder;
import com.ford.pd.bom.domain.data.manager.FB4ProgramConfigurationDataHolder;
import com.ford.pd.bom.ui.cmf.model.BCMModel;
import com.ford.pd.bom.ui.cmf.model.BomImpactModel;
import com.ford.pd.bom.ui.cmf.model.ImpactAssementCompareModel;
import com.ford.pd.bom.ui.cmf.model.NonBomImpactModel;
import com.ford.pd.bom.ui.cmf.util.CmfUtils;
import com.ford.pd.bom.ui.common.services.BCMUIService;
import com.google.common.base.Strings;

import edu.emory.mathcs.backport.java.util.Collections;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableColumn.CellDataFeatures;
import javafx.scene.control.TableView;
import javafx.util.Callback;

public class FillInTemplate extends Task<Void> {
    // Constants start

    private static final String STATUS = "CMF Status";
    private static final String TITLE = "Change Request Title";
    private static final String CR_ID = "CMF Number";
    private static final String AUTHOR = "Author CDSID";
    private static final String CHAMPION = "Other Team Members";
    private static final String COMP_DATE = "CMF Completion Date";
    private static final String REQD_DATE = "Decision Required Date";
    private static final String EXP_CHANGE = "Explanation of Change:";
    private static final String CURR_STATE = "Current State / Problem Description:";
    private static final String PME = "Principle ME/GS2";
    private static final String VL_AFFECTED = "Other VL(s) Affected";
    private static final String UNAFFECTED_PROG = "Unaffected Programs (with common parts)";
    private static final String PMT = "PMT(s) Affected:";
    private static final String TIME_STAMP_PLACE_HOLDER = "<TIME>";
    private static final String LINK_PLACE_HOLDER = "<LINK>";
    private static final String LINK = "https://it1.spt.ford.com/sites/BOMinTC/Docs/Change%20Management/Program_Mgt_Feedback/";
    private static final String FUNDING_SRC = "Funding Source";
    private static final String PRE_TOOLING_COL = "Pre-Change Tooling (Absolute)";
    private static final String POST_TOOLING_COL = "Post-Change Tooling (Absolute)";
    private static final String PRE_COST_COL = "Pre-Change Material Cost (Absolute)";
    private static final String POST_COST_COL = "Post-Change Material Cost (Absolute)";
    private static final String DELTA_COST_COL = "Incremental Change (Delta Cost)";
    private static final String DELTA_TOOLING_COL = "Incremental Change (Delta Tooling)";
    private static final String TAKE_RATE = "Take Rate";
    private static final String TAKE_RATE_OLD_COL = "Take Rate Old";
    private static final String TAKE_RATE_NEW_COL = "Take Rate New";
    private static final String QTY_OLD_COL = "Quantity Old";
    private static final String QTY_NEW_COL = "Quantity New";
    private static final String AVG_IMPACT_COL = "Avg Impact";
    private static final String FIRST_ROW = "First";
    private static final String MIDDLE_ROWS = "Odd";
    private static final String LAST_ROW = "Last";
    private static final String ONLY_ROW = "Only Row";

    private static final int EXCEL_CR_ATTR_COUNT = 14;
    private static final int OPTION_ROW_START = 17;
    private static final int CUSTOM_FIELDS_COLUMN_WIDTH = 6000;
    private static final int FINANCIAL_DETAILS_ROW_START = 5;
    private static final int CMF_SUMMARY_SHEET_POSITION = 0;
    private static final int FINANCIAL_DETAILS_SHEET_POSITION = 1;
    private static final int OTHER_IMPACTS_SHEET_POSITION = 2;
    private static final int PART_DETAIL_SHEET_POSITION = 3;

    private Desktop desktop = null;

    private String strNewTemplateFilePath = "";

    private List<BomImpactModel> relatedBomModels;

    private List<ImpactAssementCompareModel> compareBomModels;

    private List<NonBomImpactModel> nonBomModels;

    private Map<String, Map<String, String>> customFieldsModel;

    private String principalGs2;

    private CRDetailModel crModel;

    private String pmtAffected = CmfConstants.NOT_APPLICABLE;

    private String directMangedEvents = CmfConstants.NOT_APPLICABLE;

    private String potentialMangedEvents = CmfConstants.NOT_APPLICABLE;
    private List<BCMModel> bcmModelList = new ArrayList<BCMModel>();
    private String gs2ForCR;

    private long exportedTimeStamp;

    private final Map<String, Map<String, String>> headerOptionDescMap = new HashMap<>();

    private TableView<BCMModel> bcmExcelTable = null;

    private XSSFColor blueColor;
    private XSSFColor yellowColor;
    private XSSFColor lightPinkColor;

    private XSSFDataFormat formatter = null;

    public FillInTemplate(final CRDetailModel crModel, final String principalGs2) {
        super();
        this.exportedTimeStamp = System.currentTimeMillis();
        this.crModel = crModel;
        this.principalGs2 = principalGs2;
        init();
    }

    private void init() {
        this.relatedBomModels = CmfUtils.allOptionsRelatedBomImpact(this.crModel);
        this.compareBomModels = CmfUtils.getCompareAllOptionsForExport(this.crModel);
        this.nonBomModels = CmfUtils.populateExportToExcelForAllOptionsNonBOM(this.crModel);
        this.customFieldsModel = this.crModel.getCustomFieldsCR();
        populatePMTAndManagedEventsAffected();
        loadCustomFieldsHeaderDescriptionData();
        populateBCMModels();
        createCellColors();
        this.desktop = Desktop.getDesktop();
        final String title = CmfUtils.exportToExcelTitle(getChangeRequest().getChangeRequestId(),
                "ChangeSummary") + ".xlsx";
        try {
            copyNewTemplateFileFromResources(title, BomUIPanelUtil.CMF_SUMMARY_OUTLINE_TEMPLATE_PATH);
            this.strNewTemplateFilePath = System.getProperty("user.home") + File.separator + "\\Downloads\\"
                                          + title;
        } catch (final IOException e1) {
            e1.printStackTrace();
        }
    }

    private FB4ChangeRequest getChangeRequest() {
        return this.crModel.getCr();
    }

    /**
     * Populates PMT's and Managed Events affected by the CR
     */
    private void populatePMTAndManagedEventsAffected() {
        final List<String> pmts = new ArrayList<>();
        /*
         * Get CWT Impacts and get all the PMT(s) associated with it.
         */
        for (final List<Tuple<ImpactAssesmentVO, List<ImpactAssesmentVO>>> list : this.crModel.getBomImpactWithDetail().values()) {
            for (final Tuple<ImpactAssesmentVO, List<ImpactAssesmentVO>> tuple : list) {
                final String pmt = BomUtil.getFormattedPMT(tuple.x.getPmt()).replace(BomSearchConstants.PMT, "");
                if (!pmts.contains(pmt)) {
                    pmts.add(pmt);
                }
            }
        }
        if (!pmts.isEmpty()) {
            Collections.sort(pmts);
            this.pmtAffected = StringUtils.join(pmts, BomSearchConstants.SEPERATOR);
        }
        String bomWorkPuid = null;
        for (final FB4ChangeGroup cg : getChangeRequest().getChangeGroupRef()) {
            if (CmfConstants.ZERO.equals(cg.getChangeGroupID())) {
                bomWorkPuid = cg.getUniqueId();
                break;
            }
        }
        final List<String> directMEs = new ArrayList<>();
        final List<String> potentialMEs = new ArrayList<>();
        /*
         * Get all the Managed Events associated with the CR including Direct and Potential from BOM Work
         */
        final FB4ProgramConfigurationDataHolder holder = FB4ProgramConfigurationDataHolder.getInstance();
        for (final Entry<String, List<ImpactAssesmentVO>> entry : this.crModel.getNonBomImpact().entrySet()) {
            for (final ImpactAssesmentVO vo : entry.getValue()) {
                if (Strings.isNullOrEmpty(vo.getRingType()) || ImpactAssessmentConstants.DIRECT.equals(vo.getRingType())) {
                    final FB4ManagedEvent meRef = holder.getManagedEvent(vo.getManagedEvent().getUniqueId());
                    if (meRef != null && !directMEs.contains(meRef.getManagedEventName())) {
                        directMEs.add(meRef.getManagedEventName());
                    }
                }
                if (bomWorkPuid != null && entry.getKey().equals(bomWorkPuid)) {
                    if (ImpactAssessmentConstants.POTENTIAL.equals(vo.getRingType())) {
                        final FB4ManagedEvent meRef = holder.getManagedEvent(vo.getManagedEvent().getUniqueId());
                        if (meRef != null && !potentialMEs.contains(meRef.getManagedEventName())) {
                            potentialMEs.add(meRef.getManagedEventName());
                        }
                    }
                }
            }
        }
        if (!directMEs.isEmpty())
            this.directMangedEvents = StringUtils.join(directMEs, BomSearchConstants.SEPERATOR);
        if (!potentialMEs.isEmpty())
            this.potentialMangedEvents = StringUtils.join(potentialMEs, BomSearchConstants.SEPERATOR);
        this.gs2ForCR = CmfUtils.getGs2ForCR(this.crModel);
        if (Strings.isNullOrEmpty(this.gs2ForCR)) {
            this.gs2ForCR = CmfConstants.NOT_APPLICABLE;
        }
    }

    private void populateBCMModels() {
        final BCMUIService service = new BCMUIService(this.crModel);
        service.setUserCDSID("");
        service.populateBCMModelList();
        this.bcmModelList = service.getBcmModelList();
    }

    public void createExcel() {
        try {
            final int tmpRows = 22;
            final FileInputStream fileInStream = new FileInputStream(this.strNewTemplateFilePath);
            final XSSFWorkbook workbook = new XSSFWorkbook(fileInStream);
            this.formatter = workbook.createDataFormat();
            final List<String> totalOptions = new ArrayList<>();
            for (int count = 0; count < this.compareBomModels.size(); count++) {
                final ImpactAssementCompareModel model = this.compareBomModels.get(count);
                if (!totalOptions.contains(model.getOptionNumber())) {
                    totalOptions.add(model.getOptionNumber());
                }
            }
            final List<String> current3Options = new ArrayList<>();
            int CurProcess = 0;
            for (int NewSheetsCounter = 0; NewSheetsCounter < totalOptions.size(); NewSheetsCounter++) {
                current3Options.add(totalOptions.get(NewSheetsCounter));
                CurProcess++;
                updateProgress(CurProcess, totalOptions.size());
                if (current3Options.size() == 3 || NewSheetsCounter == totalOptions.size() - 1) {
                    int StartRowAt = 21;
                    int TemplateRowOptionAt = 53;
                    final XSSFSheet sheet = workbook.getSheetAt(0);
                    sheet.setDefaultRowHeightInPoints(100);
                    sheet.setAutobreaks(false);
                    sheet.setRowBreak(62);
                    for (int intOptionsCounter = 0; intOptionsCounter < current3Options.size(); intOptionsCounter++) {
                        final String curOption = current3Options.get(intOptionsCounter);
                        int intFirstOptionRow = 0;
                        final ArrayList<Integer> sameOptionsLineNumbers = new ArrayList<Integer>();
                        for (int tableCounter = 0; tableCounter < this.compareBomModels.size(); tableCounter++) {
                            final ImpactAssementCompareModel compareModel = this.compareBomModels.get(tableCounter);
                            if (curOption.equals(compareModel.getOptionNumber())) {
                                if (ImpactAssessmentConstants.COST.equals(compareModel.getRollupType())
                                    || ImpactAssessmentConstants.TOOLING.equals(compareModel.getRollupType())
                                    || ImpactAssessmentConstants.WEIGHT.equals(compareModel.getRollupType())) {
                                    final int foundRow = checkRowAlreadyExist(sameOptionsLineNumbers, sheet, compareModel);
                                    if (foundRow != -1) {
                                        final Row _xWorkingRow = sheet.getRow(foundRow);
                                        final Cell _xProgCell = _xWorkingRow.getCell(4);
                                        final Cell _xCMCell = _xWorkingRow.getCell(6);
                                        final Cell _xCostCellValue = _xWorkingRow.getCell(8);
                                        final Cell _xWeightCellValue = _xWorkingRow.getCell(10);
                                        final Cell _xToolingCellValue = _xWorkingRow.getCell(12);
                                        String value = "";
                                        if (CmfConstants.MASK.equals(compareModel.getAction())) {
                                            _xProgCell.setCellValue(CmfConstants.HIDDEN_TEXT);
                                            _xCMCell.setCellValue(CmfConstants.HIDDEN_TEXT);
                                            value = CmfConstants.HIDDEN_TEXT;
                                        } else {
                                            _xProgCell.setCellValue(compareModel.getManagedEventName());
                                            _xCMCell.setCellValue(compareModel.getControlModelName());
                                            value = compareModel.getAlt1Amount().toString();
                                        }
                                        if (ImpactAssessmentConstants.COST.equals(compareModel.getRollupType())) {
                                            _xCostCellValue.setCellValue(value);
                                        } else if (ImpactAssessmentConstants.WEIGHT.equals(compareModel.getRollupType())) {
                                            _xWeightCellValue.setCellValue(value);
                                        } else if (ImpactAssessmentConstants.TOOLING.equals(compareModel.getRollupType())) {
                                            _xToolingCellValue.setCellValue(value);
                                        }
                                        final Cell _xEDTCellValue = _xWorkingRow.getCell(15);
                                        _xEDTCellValue.setCellValue(returnAttributeValue(compareModel.getControlModelName(),
                                                compareModel.getManagedEventName(), compareModel.getOptionNumber(),
                                                CmfConstants.EDT));

                                        final Cell _xSBUVOCellValue = _xWorkingRow.getCell(18);
                                        _xSBUVOCellValue.setCellValue(returnAttributeValue(compareModel.getControlModelName(),
                                                compareModel.getManagedEventName(), compareModel.getOptionNumber(),
                                                CmfConstants.SBUVO));

                                        final Cell _xPrototypeTCellValue = _xWorkingRow.getCell(21);
                                        _xPrototypeTCellValue.setCellValue(returnAttributeValue(compareModel.getControlModelName(),
                                                compareModel.getManagedEventName(), compareModel.getOptionNumber(),
                                                CmfConstants.PROTOTOOLING));

                                        final Cell _xHEADSCellValue = _xWorkingRow.getCell(27);
                                        _xHEADSCellValue.setCellValue(returnAttributeValue(compareModel.getControlModelName(),
                                                compareModel.getManagedEventName(), compareModel.getOptionNumber(),
                                                CmfConstants.NOOFHEADS));

                                        final Cell _xAverageRevenueCellValue = _xWorkingRow.getCell(31);
                                        _xAverageRevenueCellValue.setCellValue(returnAttributeValue(
                                                compareModel.getControlModelName(), compareModel.getManagedEventName(),
                                                compareModel.getOptionNumber(), CmfConstants.REVENUE));
                                        final Cell _xTotalInvCellValue = _xWorkingRow.getCell(24);
                                        if (CmfConstants.MASK.equals(compareModel.getAction())) {
                                            _xTotalInvCellValue.setCellValue(CmfConstants.HIDDEN_TEXT);
                                        } else {
                                            final Cell _xToolingCellValue1 = _xWorkingRow.getCell(12);
                                            switch (_xToolingCellValue.getCellTypeEnum()) {
                                            case STRING:
                                                _xTotalInvCellValue.setCellValue(
                                                        getDoubleOrZeroValueFromTable(_xToolingCellValue1.getStringCellValue()) +
                                                                                 getDoubleOrZeroValueFromTable(returnAttributeValue(
                                                                                         compareModel.getControlModelName(),
                                                                                         compareModel.getManagedEventName(),
                                                                                         compareModel.getOptionNumber(),
                                                                                         CmfConstants.SBUVO))
                                                                                 +
                                                                                 getDoubleOrZeroValueFromTable(returnAttributeValue(
                                                                                         compareModel.getControlModelName(),
                                                                                         compareModel.getManagedEventName(),
                                                                                         compareModel.getOptionNumber(),
                                                                                         CmfConstants.PROTOTOOLING)));
                                                break;
                                            default:
                                                _xTotalInvCellValue.setCellValue(
                                                        _xToolingCellValue.getNumericCellValue() +
                                                                                 getDoubleOrZeroValueFromTable(returnAttributeValue(
                                                                                         compareModel.getControlModelName(),
                                                                                         compareModel.getManagedEventName(),
                                                                                         compareModel.getOptionNumber(),
                                                                                         CmfConstants.SBUVO))
                                                                                 +
                                                                                 getDoubleOrZeroValueFromTable(returnAttributeValue(
                                                                                         compareModel.getControlModelName(),
                                                                                         compareModel.getManagedEventName(),
                                                                                         compareModel.getOptionNumber(),
                                                                                         CmfConstants.PROTOTOOLING)));
                                                break;
                                            }
                                        }
                                    } else {
                                        intFirstOptionRow = intFirstOptionRow + 1;
                                        StartRowAt = StartRowAt + 1;
                                        copyRow(workbook, sheet, TemplateRowOptionAt, StartRowAt, 1);
                                        TemplateRowOptionAt = TemplateRowOptionAt + 1;
                                        final Row _xWorkingRow = sheet.getRow(StartRowAt);
                                        if (intFirstOptionRow == 1) {
                                            sameOptionsLineNumbers.add(StartRowAt);
                                            final Cell _xOptionColCell = _xWorkingRow.getCell(2);
                                            if (CmfConstants.ZERO.equals(curOption)) {
                                                _xOptionColCell.setCellValue("Bom Work");
                                            } else
                                                _xOptionColCell.setCellValue("Option #" + curOption);
                                            final CellStyle CellAllignmentCenter = _xOptionColCell.getCellStyle();
                                            CellAllignmentCenter.setAlignment(HorizontalAlignment.CENTER);
                                            _xOptionColCell.setCellStyle(CellAllignmentCenter);
                                            final CellStyle backgroundStyle = workbook.createCellStyle();
                                            backgroundStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
                                            backgroundStyle.setBorderBottom(BorderStyle.MEDIUM);
                                            backgroundStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                                            backgroundStyle.setBorderLeft(BorderStyle.MEDIUM);
                                            backgroundStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                                            backgroundStyle.setBorderRight(BorderStyle.MEDIUM);
                                            backgroundStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
                                            backgroundStyle.setBorderTop(BorderStyle.MEDIUM);
                                            backgroundStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
                                            backgroundStyle.setAlignment(HorizontalAlignment.CENTER);
                                            backgroundStyle.setWrapText(true);
                                            backgroundStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                                            backgroundStyle.setAlignment(HorizontalAlignment.CENTER);
                                            _xOptionColCell.setCellStyle(backgroundStyle);
                                        }
                                        int FoundLocation = 0;
                                        for (int i = StartRowAt; i >= tmpRows; i--) {
                                            final Row _xMergeCellsWorkingRow = sheet.getRow(i);
                                            try {
                                                if (!_xMergeCellsWorkingRow.getCell(2).getStringCellValue().equals("")) {
                                                    FoundLocation = i;
                                                    break;
                                                }
                                            } catch (final Exception e) {
                                                e.printStackTrace();
                                            }
                                        }
                                        if ((FoundLocation > 0) && (StartRowAt >= FoundLocation)) {
                                            try {
                                                getNbOfMergedRegions(sheet, FoundLocation);
                                                sheet.addMergedRegion(new CellRangeAddress(FoundLocation, StartRowAt, 2, 3));
                                            } catch (final Exception e) {
                                                System.out.println(e.getMessage());
                                            }
                                        }
                                        final Cell _xProgCell = _xWorkingRow.getCell(4);
                                        final Cell _xCMCell = _xWorkingRow.getCell(6);
                                        final Cell _xCostCellValue = _xWorkingRow.getCell(8);
                                        final Cell _xWeightCellValue = _xWorkingRow.getCell(10);
                                        final Cell _xToolingCellValue = _xWorkingRow.getCell(12);
                                        /*
                                         *If the action is MASK show the hidden text
                                         */
                                        String value = "";
                                        if (CmfConstants.MASK.equals(compareModel.getAction())) {
                                            _xProgCell.setCellValue(CmfConstants.HIDDEN_TEXT);
                                            _xCMCell.setCellValue(CmfConstants.HIDDEN_TEXT);
                                            value = CmfConstants.HIDDEN_TEXT;
                                        } else {
                                            _xProgCell.setCellValue(compareModel.getManagedEventName());
                                            _xCMCell.setCellValue(compareModel.getControlModelName());
                                            value = compareModel.getAlt1Amount().toString();
                                        }
                                        if (ImpactAssessmentConstants.COST.equals(compareModel.getRollupType())) {
                                            _xCostCellValue.setCellValue(value);
                                        } else if (ImpactAssessmentConstants.WEIGHT.equals(compareModel.getRollupType())) {
                                            _xWeightCellValue.setCellValue(value);
                                        } else if (ImpactAssessmentConstants.TOOLING.equals(compareModel.getRollupType())) {
                                            _xToolingCellValue.setCellValue(value);
                                        }

                                        final Cell _xEDTCellValue = _xWorkingRow.getCell(15);
                                        _xEDTCellValue.setCellValue(returnAttributeValue(compareModel.getControlModelName(),
                                                compareModel.getManagedEventName(), compareModel.getOptionNumber(),
                                                CmfConstants.EDT));

                                        final Cell _xSBUVOCellValue = _xWorkingRow.getCell(18);
                                        _xSBUVOCellValue.setCellValue(returnAttributeValue(compareModel.getControlModelName(),
                                                compareModel.getManagedEventName(), compareModel.getOptionNumber(),
                                                CmfConstants.SBUVO));

                                        final Cell _xPrototypeTCellValue = _xWorkingRow.getCell(21);
                                        _xPrototypeTCellValue.setCellValue(returnAttributeValue(compareModel.getControlModelName(),
                                                compareModel.getManagedEventName(), compareModel.getOptionNumber(),
                                                CmfConstants.PROTOTOOLING));

                                        final Cell _xHEADSCellValue = _xWorkingRow.getCell(27);
                                        _xHEADSCellValue.setCellValue(returnAttributeValue(compareModel.getControlModelName(),
                                                compareModel.getManagedEventName(), compareModel.getOptionNumber(),
                                                CmfConstants.NOOFHEADS));

                                        final Cell _xAverageRevenueCellValue = _xWorkingRow.getCell(31);
                                        _xAverageRevenueCellValue.setCellValue(returnAttributeValue(
                                                compareModel.getControlModelName(), compareModel.getManagedEventName(),
                                                compareModel.getOptionNumber(), CmfConstants.REVENUE));

                                        final Cell _xTotalInvCellValue = _xWorkingRow.getCell(24);
                                        /*
                                         *If the action is MASK show the hidden text
                                         */
                                        if (CmfConstants.MASK.equals(compareModel.getAction())) {
                                            _xTotalInvCellValue.setCellValue(CmfConstants.HIDDEN_TEXT);
                                        } else {
                                            final Cell _xToolingCellValue1 = _xWorkingRow.getCell(12);
                                            switch (_xToolingCellValue.getCellTypeEnum()) {
                                            case STRING:
                                                _xTotalInvCellValue.setCellValue(
                                                        getDoubleOrZeroValueFromTable(_xToolingCellValue1.getStringCellValue()) +
                                                                                 getDoubleOrZeroValueFromTable(returnAttributeValue(
                                                                                         compareModel.getControlModelName(),
                                                                                         compareModel.getManagedEventName(),
                                                                                         compareModel.getOptionNumber(),
                                                                                         CmfConstants.SBUVO))
                                                                                 +
                                                                                 getDoubleOrZeroValueFromTable(returnAttributeValue(
                                                                                         compareModel.getControlModelName(),
                                                                                         compareModel.getManagedEventName(),
                                                                                         compareModel.getOptionNumber(),
                                                                                         CmfConstants.PROTOTOOLING)));
                                                break;
                                            default:
                                                _xTotalInvCellValue.setCellValue(
                                                        _xToolingCellValue.getNumericCellValue() +
                                                                                 getDoubleOrZeroValueFromTable(returnAttributeValue(
                                                                                         compareModel.getControlModelName(),
                                                                                         compareModel.getManagedEventName(),
                                                                                         compareModel.getOptionNumber(),
                                                                                         CmfConstants.SBUVO))
                                                                                 +
                                                                                 getDoubleOrZeroValueFromTable(returnAttributeValue(
                                                                                         compareModel.getControlModelName(),
                                                                                         compareModel.getManagedEventName(),
                                                                                         compareModel.getOptionNumber(),
                                                                                         CmfConstants.PROTOTOOLING)));
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    fileInStream.close();
                    sheet.autoSizeColumn(2);
                    final int NewBottomSectionLocation = ((StartRowAt - (tmpRows)) + 34);
                    loadData(current3Options, workbook, sheet, ((StartRowAt + 1 - (tmpRows)) + 66),
                            ((StartRowAt + 1 - (tmpRows)) + 67), ((StartRowAt + 1 - (tmpRows)) + 64), NewBottomSectionLocation);
                    sheet.setDefaultRowHeightInPoints(100);
                    sheet.setDefaultRowHeightInPoints(100);
                    sheet.setFitToPage(true);
                    current3Options.clear();
                }
            }
            final XSSFSheet finalSheet = workbook.getSheetAt(0);
            /*
             * If there are no options available delete all the unused rows.
             */
            if (totalOptions.isEmpty()) {
                final int lastRowNum = finalSheet.getLastRowNum();
                for (int r = 0; r <= lastRowNum; r++) {
                    if (r >= OPTION_ROW_START) {
                        final Row row = finalSheet.getRow(r);
                        if (row != null) {
                            finalSheet.removeRow(row);
                        }
                    }
                }
            }
            /*
             * Get All the physical rows and populate specific cell values based on text.
             */
            int count = 0;
            for (int r = 0; r < finalSheet.getPhysicalNumberOfRows(); r++) {
                final XSSFRow row = finalSheet.getRow(r);
                if (row != null) {
                    for (int c = 0; c < row.getPhysicalNumberOfCells(); c++) {
                        final XSSFCell cell = row.getCell(c);
                        if (cell != null) {
                            if (CellType.STRING.equals(cell.getCellTypeEnum())) {
                                final String cellKey = cell.getStringCellValue();
                                String nextCellVal = "";
                                Cell nextCell = null;
                                Integer range = null;
                                if (CR_ID.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    nextCellVal = getChangeRequest().getChangeRequestId();
                                } else if (TITLE.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    nextCellVal = getChangeRequest().getChangeRequestTitle();
                                } else if (AUTHOR.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    nextCellVal = getChangeRequest().getAuthorCds();
                                } else if (CHAMPION.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    if (CollectionUtils.isNotEmpty(getChangeRequest().getChampionCDSIds())) {
                                        nextCellVal = StringUtils.join(getChangeRequest().getChampionCDSIds(), ",");
                                    }
                                } else if (STATUS.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    nextCellVal = getChangeRequest().getState();
                                } else if (COMP_DATE.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    if (getChangeRequest().getDecisionCompletedDate() != 0) {
                                        nextCellVal = BomUtil.longToString(getChangeRequest().getDecisionCompletedDate(),
                                                "MMM-dd-yyyy");
                                    }
                                } else if (REQD_DATE.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    if (getChangeRequest().getDecisionReqdDate() != 0) {
                                        nextCellVal =
                                                BomUtil.longToString(getChangeRequest().getDecisionReqdDate(), "MMM-dd-yyyy");
                                    }
                                } else if (EXP_CHANGE.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    if (!Strings.isNullOrEmpty(getChangeRequest().getDecisionRequested()))
                                        nextCellVal = getChangeRequest().getDecisionRequested();
                                } else if (CURR_STATE.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    if (!Strings.isNullOrEmpty(getChangeRequest().getAssumptions()))
                                        nextCellVal = getChangeRequest().getAssumptions();
                                } else if (PME.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    nextCellVal = getPrincipalMEGS2();
                                } else if (BomSearchConstants.GS2.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    nextCellVal = this.gs2ForCR;
                                } else if (PMT.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    nextCellVal = this.pmtAffected;
                                } else if (VL_AFFECTED.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    nextCellVal = this.directMangedEvents;
                                } else if (UNAFFECTED_PROG.equals(cellKey)) {
                                    range = getCellRange(finalSheet, cell);
                                    nextCellVal = this.potentialMangedEvents;
                                }
                                if (range != null) {
                                    nextCell = row.getCell(c + range);
                                    nextCell.setCellValue(nextCellVal);
                                    if (CmfConstants.NOT_APPLICABLE.equals(nextCellVal)) {
                                        final CellStyle style = nextCell.getCellStyle();
                                        style.setVerticalAlignment(VerticalAlignment.DISTRIBUTED);
                                        nextCell.setCellStyle(style);
                                    }
                                    count++;
                                }
                            }
                        }
                    }
                }
                if (EXCEL_CR_ATTR_COUNT == count) {
                    break;
                }
            }
            workbook.setSheetName(CMF_SUMMARY_SHEET_POSITION, "CMF Summary");
            workbook.createSheet("Other Impacts");
            workbook.setSheetOrder("Other Impacts", OTHER_IMPACTS_SHEET_POSITION);
            loadCustomFieldsSheet(workbook);
            populateFinancialPartDetails(workbook);
            XSSFSheet versionSheet = null;
            /*
             * Get the Version Sheet to update the exported time
             */
            for (int i = 0; i <= workbook.getNumberOfSheets(); i++) {
                final XSSFSheet sheet = workbook.getSheetAt(i);
                if (sheet.getSheetName().contains("Version")) {
                    versionSheet = sheet;
                    break;
                }
            }
            if (versionSheet != null) {
                /*
                 * Create Cell Style, Font and Hyperlink for the Link
                 */
                final XSSFCreationHelper helper = workbook.getCreationHelper();
                final CellStyle hlink_style = workbook.createCellStyle();
                hlink_style.setWrapText(true);
                hlink_style.setAlignment(HorizontalAlignment.LEFT);
                hlink_style.setVerticalAlignment(VerticalAlignment.JUSTIFY);
                final Font hlink_font = workbook.createFont();
                hlink_font.setUnderline(Font.U_SINGLE);
                hlink_font.setColor(IndexedColors.BLUE.getIndex());
                hlink_style.setFont(hlink_font);
                final Hyperlink link = helper.createHyperlink(HyperlinkType.URL);
                link.setAddress(LINK);

                for (int r = 0; r < versionSheet.getPhysicalNumberOfRows(); r++) {
                    final XSSFRow row = versionSheet.getRow(r);
                    if (row != null) {
                        for (int c = 0; c < row.getPhysicalNumberOfCells(); c++) {
                            final XSSFCell cell = row.getCell(c);
                            if (cell != null) {
                                if (CellType.STRING.equals(cell.getCellTypeEnum())) {
                                    if (TIME_STAMP_PLACE_HOLDER.equals(cell.getStringCellValue())) {
                                        final Calendar cal = Calendar.getInstance();
                                        final TimeZone timeZone = cal.getTimeZone();
                                        timeZone.getDisplayName();
                                        final SimpleDateFormat dateFormat =
                                                new SimpleDateFormat("MMM-dd-yyyy hh:mm aaa zzz", Locale.ENGLISH);
                                        dateFormat.setTimeZone(timeZone);
                                        cell.setCellValue(dateFormat.format(new Date(this.exportedTimeStamp)));
                                    } else if (LINK_PLACE_HOLDER.equals(cell.getStringCellValue())) {
                                        cell.setCellValue(LINK);
                                        cell.setHyperlink(link);
                                        cell.setCellStyle(hlink_style);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            prepareBCMExcelTable();
            exportBCMExcel(this.bcmExcelTable, "Part Details", workbook);
            workbook.setSheetOrder("Part Details", PART_DETAIL_SHEET_POSITION);
            protectSheets(workbook);
            final FileOutputStream output_file = new FileOutputStream(this.strNewTemplateFilePath);
            final ExecutorService service = Executors.newSingleThreadExecutor();
            service.submit(new OpenTemplateFileTask(workbook, this.strNewTemplateFilePath, output_file, this.desktop));
        } catch (final FileExistsException e) {
            e.printStackTrace();
        } catch (final IOException e) {
            e.printStackTrace();
        }
    }

    private void createCellColors() {
        final byte[] blueRGB = new byte[] {(byte)141, (byte)180, (byte)226};
        this.blueColor = new XSSFColor(blueRGB, null);
        final byte[] yellowRGB = new byte[] {(byte)255, (byte)192, (byte)0};
        this.yellowColor = new XSSFColor(yellowRGB, null);
        final byte[] lightPinkRGB = new byte[] {(byte)228, (byte)223, (byte)236};
        this.lightPinkColor = new XSSFColor(lightPinkRGB, null);
    }

    private void populateFinancialPartDetails(final XSSFWorkbook workbook) {
        final Map<String, List<BCMModel>> managedEventToModel = new HashMap<>();
        final Map<Integer, String> colIndexToName = new HashMap<>();
        for (final BCMModel model : this.bcmModelList) {
            if (managedEventToModel.containsKey(model.getManagedEvent())) {
                managedEventToModel.get(model.getManagedEvent()).add(model);
            } else {
                managedEventToModel.put(model.getManagedEvent(), new ArrayList<>(Arrays.asList(model)));
            }
        }
        final XSSFSheet sheet = workbook.getSheetAt(FINANCIAL_DETAILS_SHEET_POSITION);
        sheet.getSheetName();
        final XSSFRow colHeaderRow = sheet.getRow(FINANCIAL_DETAILS_ROW_START - 1);
        int takeRateColIndex = 0;
        int fundingSrcColIndex = 0;
        int toolingColIndex = 0;
        for (int i = 0; i <= colHeaderRow.getPhysicalNumberOfCells(); i++) {
            final Cell cell = colHeaderRow.getCell(i);
            if (cell != null) {
                colIndexToName.put(cell.getColumnIndex(), cell.getStringCellValue());
                if (TAKE_RATE_OLD_COL.equals(cell.getStringCellValue())) {
                    takeRateColIndex = cell.getColumnIndex();
                } else if (FUNDING_SRC.equals(cell.getStringCellValue())) {
                    fundingSrcColIndex = cell.getColumnIndex();
                } else if (POST_TOOLING_COL.equals(cell.getStringCellValue())) {
                    toolingColIndex = cell.getColumnIndex();
                }
            }
        }
        final Map<String, String> excelColNamesToFields = BCMExcelConstants.EXCEL_COL_NAMES_TO_FIELDS;
        int rowStart = FINANCIAL_DETAILS_ROW_START;

        final XSSFCellStyle yellowColorStyle = createBorderStyle(workbook);
        yellowColorStyle.setFillForegroundColor(this.yellowColor);
        yellowColorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        final XSSFCellStyle avgCellStyle = createBorderStyle(workbook);
        avgCellStyle.setDataFormat(this.formatter.getFormat("$0.00"));

        final XSSFCellStyle toolingCellStyle = createBorderStyle(workbook);
        toolingCellStyle.setDataFormat(this.formatter.getFormat("$#,##0"));

        final XSSFCellStyle mergedCellStyle = workbook.createCellStyle();
        mergedCellStyle.setAlignment(HorizontalAlignment.LEFT);
        mergedCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        final XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFamily(FontFamily.MODERN);
        font.setFontHeight(12);
        mergedCellStyle.setFont(font);
        mergedCellStyle.setFillForegroundColor(this.lightPinkColor);
        mergedCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        final XSSFCellStyle managedEventStyle = workbook.createCellStyle();
        final XSSFFont textFont = workbook.createFont();
        textFont.setColor(this.lightPinkColor);
        managedEventStyle.setFont(textFont);
        managedEventStyle.setFillForegroundColor(this.lightPinkColor);
        managedEventStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        String avgStartCell = "";
        String avgEndCell = "";
        String newTakeRateCell = "";
        String newCostStartCell = "";
        String newQtyStartCell = "";
        String oldTakeRateCell = "";
        String oldCostStartCell = "";
        String oldQtyStartCell = "";

        final Map<String, CellStyle> standardCellStyleMap = createBorderStyle(workbook, "");
        final Map<String, CellStyle> takeRateCellStyleMap = createBorderStyle(workbook, TAKE_RATE);
        final Map<String, CellStyle> avgImpactCellStyleMap = createBorderStyle(workbook, AVG_IMPACT_COL);
        for (final Entry<String, List<BCMModel>> entry : managedEventToModel.entrySet()) {
            final XSSFRow mergedRow = sheet.createRow(rowStart);
            mergedRow.setHeight((short)350);
            sheet.addMergedRegion(new CellRangeAddress(rowStart, rowStart, 1, colIndexToName.size()));
            final Cell mergedCell = mergedRow.createCell(1);
            sheet.autoSizeColumn(mergedCell.getColumnIndex());
            mergedCell.setCellStyle(mergedCellStyle);
            mergedCell.setCellType(CellType.STRING);
            mergedCell.setCellValue(entry.getKey());
            rowStart++;
            int rowsAdded = 0;
            final int totalRows = entry.getValue().size();
            avgStartCell = "";
            avgEndCell = "";
            newTakeRateCell = "";
            newCostStartCell = "";
            newQtyStartCell = "";
            oldTakeRateCell = "";
            oldCostStartCell = "";
            oldQtyStartCell = "";
            double toolingSum = 0;
            for (final BCMModel model : entry.getValue()) {
                final XSSFRow row = sheet.createRow(rowStart);
                // row.setHeight((short)300);
                rowStart++;
                rowsAdded++;
                for (int i = 1; i < colIndexToName.size() + 1; i++) {
                    final Cell cell = row.createCell(i);
                    final String colName = colIndexToName.getOrDefault(cell.getColumnIndex(), "");
                    if (rowsAdded == 1) {
                        if (AVG_IMPACT_COL.equals(colName)) {
                            avgStartCell = cell.getAddress().formatAsString();
                        }
                    }
                    if (totalRows == rowsAdded) {
                        if (AVG_IMPACT_COL.equals(colName)) {
                            avgEndCell = cell.getAddress().formatAsString();
                        }
                    }
                    if (TAKE_RATE_OLD_COL.equals(colName)) {
                        oldTakeRateCell = cell.getAddress().formatAsString();
                    } else if (TAKE_RATE_NEW_COL.equals(colName)) {
                        newTakeRateCell = cell.getAddress().formatAsString();
                    } else if (QTY_OLD_COL.equals(colName)) {
                        oldQtyStartCell = cell.getAddress().formatAsString();
                    } else if (QTY_NEW_COL.equals(colName)) {
                        newQtyStartCell = cell.getAddress().formatAsString();
                    } else if (PRE_COST_COL.equals(colName)) {
                        oldCostStartCell = cell.getAddress().formatAsString();
                    } else if (POST_COST_COL.equals(colName)) {
                        newCostStartCell = cell.getAddress().formatAsString();
                    }
                    Object object = null;
                    String value = "";
                    try {
                        if (!Strings.isNullOrEmpty(colName)
                            && excelColNamesToFields.containsKey(colName)) {
                            object = PropertyUtils.getProperty(model, excelColNamesToFields.get(colName));
                        }
                    } catch (final Exception e) {
                        e.printStackTrace();
                    }
                    if (object != null) {
                        value = (String)object;
                    }
                    CellType cellType = null;
                    if (colName.contains(BomAuthoringConstants.QUANTITY)) {
                        value = getFormatNumber(value, BomAuthoringConstants.QUANTITY);
                        cellType = CellType.NUMERIC;
                    } else if (colName.contains(ImpactAssessmentConstants.COST)) {
                        cellType = CellType.NUMERIC;
                        if (PRE_COST_COL.equals(colName)
                            || POST_COST_COL.equals(colName)
                            || DELTA_COST_COL.equals(colName)) {
                            value = getFormatNumber(value, ImpactAssessmentConstants.COST);
                        }
                    } else if (colName.contains(ImpactAssessmentConstants.TOOLING)) {
                        cellType = CellType.NUMERIC;
                        if (PRE_TOOLING_COL.equals(colName)
                            || POST_TOOLING_COL.equals(colName)
                            || DELTA_TOOLING_COL.equals(colName)) {
                            value = getFormatNumber(value, ImpactAssessmentConstants.TOOLING);
                        }
                        if (DELTA_TOOLING_COL.equals(colName)) {
                            toolingSum = toolingSum + Double.parseDouble(CmfWorkFlowUtil.omitSpecialCharacters(value));
                        }
                    } else {
                        if (BomAuthoringConstants.CURRENCY_PROPERTY_VALUE.equalsIgnoreCase(colName)) {
                            value = "USD";
                        }
                        if (AVG_IMPACT_COL.equals(colName)) {
                            cell.setCellType(CellType.FORMULA);
                            cellType = CellType.FORMULA;
                            cell.setCellFormula("+(" + newTakeRateCell + "*" + newCostStartCell + "*" + newQtyStartCell + ")-("
                                                + oldTakeRateCell + "*" + oldCostStartCell + "*" + oldQtyStartCell + ")");
                        } else {
                            cellType = CellType.STRING;
                        }
                    }
                    if (colName.contains(TAKE_RATE)) {
                        cell.setCellType(CellType.NUMERIC);
                        value = "0%";
                        if (totalRows == 1) {
                            cell.setCellStyle(takeRateCellStyleMap.get(ONLY_ROW));
                        } else if (rowsAdded == 1) {
                            cell.setCellStyle(takeRateCellStyleMap.get(FIRST_ROW));
                        } else if (totalRows == rowsAdded) {
                            cell.setCellStyle(takeRateCellStyleMap.get(LAST_ROW));
                        } else {
                            cell.setCellStyle(takeRateCellStyleMap.get(MIDDLE_ROWS));
                        }
                    } else if (AVG_IMPACT_COL.equals(colName)) {
                        if (totalRows == 1) {
                            cell.setCellStyle(avgImpactCellStyleMap.get(ONLY_ROW));
                        } else if (rowsAdded == 1) {
                            cell.setCellStyle(avgImpactCellStyleMap.get(FIRST_ROW));
                        } else if (totalRows == rowsAdded) {
                            cell.setCellStyle(avgImpactCellStyleMap.get(LAST_ROW));
                        } else {
                            cell.setCellStyle(avgImpactCellStyleMap.get(MIDDLE_ROWS));
                        }
                    } else {
                        if (totalRows == 1) {
                            cell.setCellStyle(standardCellStyleMap.get(ONLY_ROW));
                        } else if (rowsAdded == 1) {
                            cell.setCellStyle(standardCellStyleMap.get(FIRST_ROW));
                        } else if (totalRows == rowsAdded) {
                            cell.setCellStyle(standardCellStyleMap.get(LAST_ROW));
                        } else {
                            cell.setCellStyle(standardCellStyleMap.get(MIDDLE_ROWS));
                        }
                    }
                    cell.setCellType(cellType);
                    cell.setCellValue(value);
                }
            }
            final XSSFRow summingRow = sheet.createRow(rowStart);
            summingRow.setHeight((short)300);
            sheet.addMergedRegion(new CellRangeAddress(rowStart, rowStart, 1, takeRateColIndex));
            sheet.addMergedRegion(new CellRangeAddress(rowStart, rowStart, fundingSrcColIndex, toolingColIndex));
            final Cell summingMergedCell1 = summingRow.createCell(1);
            summingMergedCell1.setCellStyle(managedEventStyle);
            summingMergedCell1.setCellValue(entry.getKey());
            final Cell summingMergedCell2 = summingRow.createCell(fundingSrcColIndex);
            summingMergedCell2.setCellStyle(mergedCellStyle);
            final Cell avgCell = summingRow.createCell(takeRateColIndex + 1);
            avgCell.setCellStyle(yellowColorStyle);
            avgCell.setCellType(CellType.STRING);
            avgCell.setCellValue("Avg.=>");
            final Cell avgSumCell = summingRow.createCell(takeRateColIndex + 2);
            avgSumCell.setCellStyle(avgCellStyle);
            avgSumCell.setCellType(CellType.FORMULA);
            avgSumCell.setCellFormula("SUM(" + avgStartCell + ":" + avgEndCell + ")");
            final Cell toolingCell = summingRow.createCell(toolingColIndex + 1);
            toolingCell.setCellStyle(yellowColorStyle);
            toolingCell.setCellType(CellType.STRING);
            toolingCell.setCellValue("Tool=>");
            final Cell toolingSumCell = summingRow.createCell(toolingColIndex + 2);
            toolingSumCell.setCellStyle(toolingCellStyle);
            toolingSumCell.setCellType(CellType.NUMERIC);
            toolingSumCell.setCellValue(toolingSum);
            rowStart++;
        }
    }

    private String getFormatNumber(final String inputVal, final String unitType) {
        String value = "0";
        if (!Strings.nullToEmpty(inputVal).trim().isEmpty()
            && !CmfConstants.NOT_APPLICABLE.equals(inputVal)) {
            value = inputVal;
        }
        if (unitType.equals(BomAuthoringConstants.QUANTITY)) {
            return CmfUtils.formatNumber(value, ImpactAssessmentConstants.TOOLING);
        }
        return CmfUtils.formatNumber(value, unitType);
    }

    private XSSFCellStyle createBorderStyle(final XSSFWorkbook workbook) {
        final XSSFCellStyle borderStyle = workbook.createCellStyle();
        borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        borderStyle.setBorderTop(BorderStyle.MEDIUM);
        borderStyle.setBorderBottom(BorderStyle.MEDIUM);
        borderStyle.setBorderLeft(BorderStyle.MEDIUM);
        borderStyle.setBorderRight(BorderStyle.MEDIUM);
        borderStyle.setAlignment(HorizontalAlignment.LEFT);
        borderStyle.setVerticalAlignment(VerticalAlignment.JUSTIFY);
        borderStyle.setWrapText(true);
        return borderStyle;
    }

    private Map<String, CellStyle> createBorderStyle(final XSSFWorkbook workbook, final String type) {
        final Map<String, CellStyle> takeRateCellStyleMap = new HashMap<>();
        final XSSFCellStyle firstRowStyle = workbook.createCellStyle();
        takeRateCellStyleMap.put(FIRST_ROW, firstRowStyle);
        final XSSFCellStyle middleRowsStyle = workbook.createCellStyle();
        takeRateCellStyleMap.put(MIDDLE_ROWS, middleRowsStyle);
        final XSSFCellStyle lastRowStyle = workbook.createCellStyle();
        takeRateCellStyleMap.put(LAST_ROW, lastRowStyle);
        final XSSFCellStyle onlyRowStyle = workbook.createCellStyle();
        takeRateCellStyleMap.put(ONLY_ROW, onlyRowStyle);
        applyBorderProperties(firstRowStyle, type);
        applyBorderProperties(middleRowsStyle, type);
        applyBorderProperties(lastRowStyle, type);
        applyBorderProperties(onlyRowStyle, type);
        firstRowStyle.setBorderTop(BorderStyle.MEDIUM);
        firstRowStyle.setBorderBottom(BorderStyle.THIN);
        middleRowsStyle.setBorderTop(BorderStyle.THIN);
        middleRowsStyle.setBorderBottom(BorderStyle.THIN);
        lastRowStyle.setBorderTop(BorderStyle.THIN);
        lastRowStyle.setBorderBottom(BorderStyle.MEDIUM);
        onlyRowStyle.setBorderTop(BorderStyle.MEDIUM);
        onlyRowStyle.setBorderBottom(BorderStyle.MEDIUM);
        return takeRateCellStyleMap;
    }

    private void applyBorderProperties(final XSSFCellStyle borderStyle, final String type) {
        borderStyle.setBorderLeft(BorderStyle.MEDIUM);
        borderStyle.setBorderRight(BorderStyle.MEDIUM);
        borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        borderStyle.setAlignment(HorizontalAlignment.LEFT);
        /*borderStyle.setVerticalAlignment(VerticalAlignment.DISTRIBUTED);*/
        borderStyle.setWrapText(true);
        if (TAKE_RATE.equals(type)) {
            borderStyle.setDataFormat(this.formatter.getFormat("0%"));
            borderStyle.setLocked(false);
            borderStyle.setFillForegroundColor(this.blueColor);
            borderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        } else if (AVG_IMPACT_COL.equals(type)) {
            borderStyle.setDataFormat(this.formatter.getFormat("$0.00"));
        } else if (ImpactAssessmentConstants.COST.equals(type)) {
            borderStyle.setDataFormat(this.formatter.getFormat("0.00"));
        } else if (ImpactAssessmentConstants.TOOLING.equals(type)) {
            borderStyle.setDataFormat(this.formatter.getFormat("#,##0"));
        }
    }

    /**
     * Creation of Part Details Excel sheet
     * load part data into the sheet.
     *
     * @param table
     * @param sheetName
     * @param xWorkbook
     */
    public void exportBCMExcel(final TableView<BCMModel> table, final String sheetName, final Workbook xWorkbook) {
        final int numberOfRows = table.getItems().size();
        final ObservableList<TableColumn<BCMModel, ?>> columns = table.getColumns();
        final Sheet xSheet = xWorkbook.createSheet(sheetName);
        final Font font = xWorkbook.createFont();
        font.setBold(true);
        int rowIndex = 0;
        final Row xRow = xSheet.createRow(rowIndex++);
        final XSSFCellStyle style = (XSSFCellStyle)xWorkbook.createCellStyle();
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setFont(font);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setRotation((short)90);

        for (int rowCount = 0; rowCount < numberOfRows; rowCount++) {
            final Row _xRow = xSheet.createRow(rowIndex++);
            for (int columCount = 0; columCount < columns.size(); columCount++) {
                final Cell _xCell = _xRow.createCell(columCount);
                if (columns.get(columCount).getCellData(rowCount) != null) {
                    final String data = columns.get(columCount).getCellData(rowCount).toString();
                    if (columns.get(columCount).getText().contains("MaterialCost - Old")
                        || columns.get(columCount).getText().contains("Material Cost - New")
                        || columns.get(columCount).getText().contains("Material Cost - Delta")) {
                        final CellStyle style2 = xWorkbook.createCellStyle();
                        style2.setDataFormat(xWorkbook.createDataFormat().getFormat("0.00"));
                        if (NumberIsDouble(data))
                            _xCell.setCellValue(Double.parseDouble(data));
                        else if (NumberIsInteger(data))
                            _xCell.setCellValue(Integer.parseInt(data));
                        else
                            _xCell.setCellValue(data);
                        _xCell.setCellStyle(style2);
                    } else if (columns.get(columCount).getText().contains("Investment - Old")
                               || columns.get(columCount).getText().contains("Investment - New")
                               || columns.get(columCount).getText().contains("Investment - Delta")) {
                        final CellStyle style2 = xWorkbook.createCellStyle();
                        style2.setDataFormat(xWorkbook.createDataFormat().getFormat("#,##0"));
                        if (NumberIsDouble(data))
                            _xCell.setCellValue(Double.parseDouble(data));
                        else if (NumberIsInteger(data))
                            _xCell.setCellValue(Integer.parseInt(data));
                        else {
                            _xCell.setCellValue(data);
                        }
                        _xCell.setCellStyle(style2);
                    } else {
                        final CellStyle style2 = xWorkbook.createCellStyle();
                        _xCell.setCellValue(data);
                        _xCell.setCellStyle(style2);
                    }

                } else {
                    _xCell.setCellValue("");

                }
            }
        }

        for (int i = 0; i < columns.size(); i++) {
            xSheet.autoSizeColumn(i);
            if (xSheet.getColumnWidth(i) < 1000) {
                xSheet.setColumnWidth(i, 1700);
            }
            final Cell xCell = xRow.createCell(i);
            String colName = columns.get(i).getText();
            if (Strings.isNullOrEmpty(colName)) {
                colName = columns.get(i).getId();
            }
            if (colName != null) {
                xCell.setCellValue(colName);
                xRow.getCell(i).setCellStyle(style);
            }
        }
        xSheet.createFreezePane(6, 1, 6, 1);
        xSheet.setAutoFilter(new CellRangeAddress(0, 0, 0, columns.size() - 1));
    }

    /**
     * check if the string given string is double
     */
    private boolean NumberIsDouble(final String whatNmber) {
        boolean returnValue = false;
        try {
            Double.parseDouble(whatNmber);
            returnValue = true;
        } catch (final NumberFormatException e) {
            returnValue = false;
        }
        return returnValue;

    }

    /**
     * check if the string given string is integer
     */
    private boolean NumberIsInteger(final String whatNumber) {
        boolean returnValue = false;
        try {
            Integer.parseInt(whatNumber);
            returnValue = true;
        } catch (final NumberFormatException e) {
            returnValue = false;
        }
        return returnValue;

    }

    /**
     * Map to hold header as key,
     * map of option and description as value.
     */
    private void loadCustomFieldsHeaderDescriptionData() {
        for (final FB4ChangeGroup cg : getChangeRequest().getChangeGroupRef()) {
            final Map<String, String> map = this.customFieldsModel.get(cg.getUniqueId());
            if (map != null)
                for (final Entry<String, String> entry : map.entrySet()) {
                    Map<String, String> valueMap = null;
                    /*
                     * if header not found in map, create new else retrieve.
                     */
                    if (!this.headerOptionDescMap.containsKey(entry.getKey()))
                        valueMap = new HashMap<>();
                    else
                        valueMap = this.headerOptionDescMap.get(entry.getKey());
                    final String option = "Option #" + cg.getChangeGroupID() + " " + cg.getChangeGroupName();
                    valueMap.put(option, entry.getValue());
                    this.headerOptionDescMap.put(entry.getKey(), valueMap);
                }
        }
    }

    /**
     * Protects all the sheets available in the workbook.
     *
     * @param workbook
     */
    private void protectSheets(final XSSFWorkbook workbook) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            final XSSFSheet sheet = workbook.getSheetAt(i);
            final PrintSetup ps = sheet.getPrintSetup();
            ps.setFitWidth((short)1);
            ps.setFitHeight((short)0);
            // sheet.setRowBreak(62);
            sheet.enableLocking();
            sheet.protectSheet("FedeBomLockP@$$1");
            final CTSheetProtection sheetProtection = sheet.getCTWorksheet().getSheetProtection();
            sheetProtection.setSelectLockedCells(false);
            sheetProtection.setSelectUnlockedCells(false);
            sheetProtection.setFormatCells(true);
            sheetProtection.setFormatColumns(true);
            sheetProtection.setFormatRows(true);
            sheetProtection.setInsertColumns(true);
            sheetProtection.setInsertRows(true);
            sheetProtection.setInsertHyperlinks(true);
            sheetProtection.setDeleteColumns(true);
            sheetProtection.setDeleteRows(true);
            sheetProtection.setSort(false);
            sheetProtection.setAutoFilter(false);
            sheetProtection.setPivotTables(true);
            sheetProtection.setObjects(true);
            sheetProtection.setScenarios(true);
            sheet.getCTWorksheet().setSheetProtection(sheetProtection);
        }
    }

    /**
     * Load Custom Fields(Other Impacts) Tab data
     *
     * @param workbook
     */
    private void loadCustomFieldsSheet(final XSSFWorkbook workbook) {
        XSSFSheet otherImpactsSheet = null;
        /*
         * Get the Other Impacts Sheet to update the exported time
         */
        for (int i = 0; i <= workbook.getNumberOfSheets(); i++) {
            final XSSFSheet sheet = workbook.getSheetAt(i);
            if (sheet.getSheetName().contains("Other Impacts")) {
                otherImpactsSheet = sheet;
                break;
            }
        }
        if (otherImpactsSheet != null) {
            int rowCount = 0;
            final Row row = otherImpactsSheet.createRow(rowCount);
            rowCount++;
            final Cell crIdCell = row.createCell(0);
            otherImpactsSheet.setColumnWidth(0, CUSTOM_FIELDS_COLUMN_WIDTH);
            crIdCell.setCellValue("CMF:" + this.crModel.getCr().getChangeRequestId());

            final XSSFFont font = workbook.createFont();
            font.setBold(true);
            font.setFontHeight(12);

            final CellStyle backgroundStyle = getBorderStyle(workbook);
            backgroundStyle.setFont(font);
            crIdCell.setCellStyle(backgroundStyle);
            final CellStyle descBackgroundStyle = getBorderStyle(workbook);
            final Row optionHeaderRow = otherImpactsSheet.createRow(rowCount);
            rowCount++;
            optionHeaderRow.createCell(0).setCellValue("Topics/Options");
            optionHeaderRow.getCell(0).setCellStyle(backgroundStyle);
            int columnCount = 1;
            /*
             * map to store column index and option name
             */
            final Map<Integer, String> columnOptionMap = new HashMap<>();

            /*
             * Create Option header Row - Fill yellow background for authorized option.
             */
            for (final FB4ChangeGroup cg : getChangeRequest().getChangeGroupRef()) {
                if (this.customFieldsModel.containsKey(cg.getUniqueId())) {
                    final String optionName = "Option #" + cg.getChangeGroupID() + " " + cg.getChangeGroupName();
                    columnOptionMap.put(columnCount, optionName);
                    final Cell optionHeaderCell = optionHeaderRow.createCell(columnCount);
                    optionHeaderCell.setCellValue(optionName);
                    if (CmfConstants.CHANGE_GROUP_RECOMMENDED.equals(cg.getChangeGroupState())) {
                        final CellStyle recommmendOptnStyle = getBorderStyle(workbook);
                        recommmendOptnStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                        recommmendOptnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        recommmendOptnStyle.setFont(font);
                        optionHeaderCell.setCellStyle(recommmendOptnStyle);
                    } else {
                        optionHeaderCell.setCellStyle(backgroundStyle);
                    }
                    otherImpactsSheet.autoSizeColumn(columnCount);
                    if (otherImpactsSheet.getColumnWidth(columnCount) < CUSTOM_FIELDS_COLUMN_WIDTH) {
                        otherImpactsSheet.setColumnWidth(columnCount, CUSTOM_FIELDS_COLUMN_WIDTH);
                    }
                    columnCount++;
                }
            }

            /*
             * Populate header and description
             */
            for (final Entry<String, Map<String, String>> entry : this.headerOptionDescMap.entrySet()) {
                final Row headerDescRow = otherImpactsSheet.createRow(rowCount);
                rowCount++;
                final Cell cell = headerDescRow.createCell(0);
                cell.setCellValue(entry.getKey());
                cell.setCellStyle(backgroundStyle);
                for (int i = 1; i < entry.getValue().size() + 1; i++) {
                    final Cell desc = headerDescRow.createCell(i);
                    desc.setCellStyle(descBackgroundStyle);
                    final String colName = columnOptionMap.get(i);
                    if (entry.getValue().containsKey(colName)) {
                        desc.setCellValue(entry.getValue().get(colName));
                    } else {
                        desc.setCellValue("");
                    }
                }
            }

            final Row recommendedOptnRow = otherImpactsSheet.createRow(rowCount + 5);
            final CellStyle coloredStyle = workbook.createCellStyle();
            coloredStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            coloredStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            final Cell coloredCell = recommendedOptnRow.createCell(2);
            coloredCell.setCellStyle(coloredStyle);
            final Cell recommendedOptnCell = recommendedOptnRow.createCell(3);
            recommendedOptnCell.setCellValue("RECOMMENDED OPTION");

        }
    }

    /**
     *
     * @param workbook
     * @return
     */
    private CellStyle getBorderStyle(final XSSFWorkbook workbook) {
        final CellStyle backgroundStyle = workbook.createCellStyle();
        backgroundStyle.setBorderBottom(BorderStyle.THIN);
        backgroundStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        backgroundStyle.setBorderLeft(BorderStyle.THIN);
        backgroundStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        backgroundStyle.setBorderRight(BorderStyle.THIN);
        backgroundStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        backgroundStyle.setBorderTop(BorderStyle.THIN);
        backgroundStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        backgroundStyle.setAlignment(HorizontalAlignment.LEFT);
        backgroundStyle.setVerticalAlignment(VerticalAlignment.JUSTIFY);
        backgroundStyle.setWrapText(true);
        return backgroundStyle;
    }

    /**
     * Returns Number of columns Merged.
     * Range First Column and Last Column are both inclusive.
     *
     * @param sheet
     * @param cell
     * @return Number of columns Merged
     */
    private int getCellRange(final XSSFSheet sheet, final Cell cell) {
        for (int r = 0; r < sheet.getNumMergedRegions(); r++) {
            final CellRangeAddress range = sheet.getMergedRegion(r);
            if (range.isInRange(cell)) {
                /*return range.getNumberOfCells();*/
                return (range.getLastColumn() - range.getFirstColumn()) + 1;
            }
        }
        return 0;
    }

    /**
     * @return Change Request Principle ME/GS2
     */
    private String getPrincipalMEGS2() {
        String value = CmfConstants.NOT_APPLICABLE;
        if (getChangeRequest().getPrincipalManagedEvent() != null) {
            if (CmfUtils.hasMeAccess(getChangeRequest().getPrincipalManagedEvent())) {
                value = getChangeRequest().getPrincipalManagedEvent().getManagedEventName();
            } else {
                value = CmfConstants.HIDDEN_TEXT;
            }
        } else {
            /*
             * If Text not equals to N/A then Principal GS2 is available
             */
            if (!CmfConstants.NOT_APPLICABLE.equals(this.principalGs2)) {
                final FB4GS2 gs2 = FB4BOMPrgIndependentDataHolder.getInstance().getGs2Container().get(this.principalGs2);
                if (gs2 != null) {
                    if (CmfUtils.hasGs2Access(gs2)) {
                        value = gs2.getGs2Description();
                    } else {
                        value = CmfConstants.HIDDEN_TEXT;
                    }
                }
            }
        }
        return value;
    }

    public static void getNbOfMergedRegions(final XSSFSheet worksheet, final int row) {
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            if ((worksheet.getMergedRegion(i).getFirstRow() == row) && (worksheet.getMergedRegion(i).getFirstColumn() == 2)) {
                worksheet.removeMergedRegion(i);
            }
        }
    }

    /**
     * @see javafx.concurrent.Task#call()
     */
    @Override
    protected Void call() {
        try {
            createExcel();
        } catch (final Exception e) {
            System.out.println(e.getMessage());
        }
        return null;
    }

    private double getDoubleOrZeroValueFromTable(final String strWhatValue) {
        final String Digits = "(\\p{Digit}+)";
        final String HexDigits = "(\\p{XDigit}+)";
        final String Exp = "[eE][+-]?" + Digits;
        final String fpRegex =
                ("[\\x00-\\x20]*" +
                 "[+-]?(" +
                 "NaN|" +
                 "Infinity|" +
                 "(((" + Digits + "(\\.)?(" + Digits + "?)(" + Exp + ")?)|" +
                 "(\\.(" + Digits + ")(" + Exp + ")?)|" +
                 "((" +
                 "(0[xX]" + HexDigits + "(\\.)?)|" +
                 "(0[xX]" + HexDigits + "?(\\.)" + HexDigits + ")" +
                 ")[pP][+-]?" + Digits + "))" +
                 "[fFdD]?))" +
                 "[\\x00-\\x20]*");

        if (Pattern.matches(fpRegex, strWhatValue)) {
            return Double.valueOf(strWhatValue);
        } else {
            return Double.valueOf(0);
        }
    }

    private String returnAttributeValue(final String cmName, final String meName, final String optionNumber,
            final String WhatAttribute) {
        String returnedValue = "";
        for (final BomImpactModel model : this.relatedBomModels) {
            if (model.getControlModelName().equals(cmName)
                && model.getManagedEventName().equals(meName)
                && model.getOptionNumber().equals(optionNumber)) {
                if (CmfConstants.EDT.equals(WhatAttribute)) {
                    returnedValue = model.getEdt();
                } else if (CmfConstants.NOOFHEADS.equals(WhatAttribute)) {
                    returnedValue = model.getNoofHeads();
                } else if (CmfConstants.REVENUE.equals(WhatAttribute)) {
                    returnedValue = model.getRevenue();
                } else if (CmfConstants.PROTOTOOLING.equals(WhatAttribute)) {
                    returnedValue = model.getPrototooling();
                } else if (CmfConstants.SBUVO.equals(WhatAttribute)) {
                    returnedValue = model.getSbuVo();
                } else if ("testing".equals(WhatAttribute)) {
                    returnedValue = "";
                } else if ("averagematerial".equals(WhatAttribute)) {
                    returnedValue = "";
                } else if ("profitmargin".equals(WhatAttribute)) {
                    returnedValue = "";
                }
                /*
                 * If the action is mask and value is not null - Mask It
                 */
                if (CmfConstants.MASK.equals(model.getAction())
                    && !Strings.isNullOrEmpty(returnedValue)) {
                    returnedValue = CmfConstants.HIDDEN_TEXT;
                }
                break;
            }
        }
        if (returnedValue != null) {
            return returnedValue;
        } else {
            return "N/A";
        }
    }

    private static int checkRowAlreadyExist(final ArrayList<Integer> Lines, final XSSFSheet sheet,
            final ImpactAssementCompareModel ArrTotalOptions) {
        for (int RowCheck = 0; RowCheck < Lines.size(); RowCheck++) {
            final Row _xWorkingRow = sheet.getRow(Lines.get(RowCheck));
            final Cell _xProgCell = _xWorkingRow.getCell(4);
            final Cell _xCMCell = _xWorkingRow.getCell(6);
            if (_xProgCell.getStringCellValue().equals(ArrTotalOptions.getManagedEventName())
                && (_xCMCell.getStringCellValue().equals(ArrTotalOptions.getControlModelName()))) {
                return Lines.get(RowCheck);
            }
        }
        return -1;
    }

    private int loadData(final List<String> current3Optionslist, final XSSFWorkbook workbook,
            final XSSFSheet worksheet, final int rowNumNewOption1, final int rowNumNewOption2, final int rowNewManagedEvent,
            final int startRowsAt) {
        int tempRowNumNewOption1 = rowNumNewOption1;
        int tempRowNumNewOption2 = rowNumNewOption2;
        int tempRowNewManagedEvent = rowNewManagedEvent;
        int setNextRowNumberInExcel = startRowsAt;
        int totalAddedRows = 0;
        boolean blnFirstOptionStart = false;
        if (this.nonBomModels.size() > 0) {
            for (int counterOptions = 0; counterOptions < current3Optionslist.size(); counterOptions++) {
                if (blnFirstOptionStart == false) {
                    blnFirstOptionStart = true;
                } else {
                    setNextRowNumberInExcel = setNextRowNumberInExcel + 2;
                }
                final Map<String, String> nonBomValuesMap = new HashMap<>();
                final Map<String, String> valuesSummaryColNumber = new HashMap<>();
                final Row _xRowSummaryTitle = worksheet.getRow(tempRowNumNewOption1);
                for (int indexCol = 6; indexCol < _xRowSummaryTitle.getLastCellNum(); indexCol++) {
                    final Cell _xKeytmpCell = _xRowSummaryTitle.getCell(indexCol);
                    if (_xKeytmpCell.getStringCellValue().length() > 0) {
                        nonBomValuesMap.put(_xKeytmpCell.getStringCellValue().toUpperCase(), null);
                    }
                    valuesSummaryColNumber.put(_xKeytmpCell.getStringCellValue().toUpperCase(), String.valueOf(indexCol));
                }
                copyRow(workbook, worksheet, tempRowNumNewOption1, setNextRowNumberInExcel, 1);
                worksheet.createRow(setNextRowNumberInExcel - 1);
                if (current3Optionslist.get(counterOptions).equals("0")) {
                    worksheet.getRow(setNextRowNumberInExcel).getCell(2).setCellValue("CMF Level: Bom Work");
                } else {
                    worksheet.getRow(setNextRowNumberInExcel)
                            .getCell(2)
                            .setCellValue("CMF Level: Option " + current3Optionslist.get(counterOptions));
                }
                final int nextRowNumber = setNextRowNumberInExcel + 1;
                copyRow(workbook, worksheet, tempRowNumNewOption1 + 1, nextRowNumber, 1);
                worksheet.getRow(nextRowNumber).getCell(2).setCellValue("Summary");
                tempRowNumNewOption1 = tempRowNumNewOption1 + 2;
                tempRowNumNewOption2 = tempRowNumNewOption2 + 2;
                tempRowNewManagedEvent = tempRowNewManagedEvent + 2;
                for (int exportTableCounter = 0; exportTableCounter < this.nonBomModels.size(); exportTableCounter++) {
                    if (this.nonBomModels.get(exportTableCounter).getAttributes().keySet().size() > 0) {
                        final int StartBlockManagedEventRowNum = setNextRowNumberInExcel;
                        int EndBlockManagedEventRowNum = 0;
                        boolean BlockFound = false;
                        final NonBomImpactModel currentModel = this.nonBomModels.get(exportTableCounter);
                        for (final String key : currentModel.getAttributes().keySet()) {
                            if (CmfConstants.YES.equals(currentModel.getAttributes().get(key))) {
                                if (currentModel.getOptionNumber().equals(current3Optionslist.get(counterOptions))) {
                                    copyRow(workbook, worksheet, tempRowNewManagedEvent, setNextRowNumberInExcel + 2, 1);
                                    tempRowNumNewOption1 = tempRowNumNewOption1 + 1;
                                    tempRowNumNewOption2 = tempRowNumNewOption2 + 1;
                                    tempRowNewManagedEvent = tempRowNewManagedEvent + 1;
                                    final Row _xRow = worksheet.getRow(setNextRowNumberInExcel + 2);
                                    _xRow.getCell(2);
                                    /*
                                     * Convert key to uppercase
                                     */
                                    final String valueSummaryKey = key.toUpperCase();
                                    if (nonBomValuesMap.containsKey(valueSummaryKey)) {
                                        if (CmfConstants.MASK.equals(currentModel.getAction())) {
                                            nonBomValuesMap.put(valueSummaryKey, CmfConstants.HIDDEN_TEXT);
                                        } else {
                                            nonBomValuesMap.put(valueSummaryKey, currentModel.getAttributes().get(key));
                                        }
                                    }
                                    final Cell _xKeyAttribute = _xRow.getCell(5);
                                    _xKeyAttribute.setCellValue(key);
                                    BlockFound = true;
                                    final Cell _xCommentCell = _xRow.getCell(7);
                                    if (currentModel.getComments().get(key) == null) {
                                        _xCommentCell.setCellValue("");
                                    } else {
                                        if (CmfConstants.MASK.equals(currentModel.getAction())) {
                                            _xCommentCell.setCellValue(CmfConstants.HIDDEN_TEXT);
                                        } else {
                                            _xCommentCell.setCellValue(currentModel.getComments().get(key));
                                        }
                                    }
                                    setNextRowNumberInExcel = setNextRowNumberInExcel + 1;
                                }
                            }
                        }
                        if (BlockFound == true) {
                            EndBlockManagedEventRowNum = setNextRowNumberInExcel;
                            final Cell cell = worksheet.getRow(StartBlockManagedEventRowNum).getCell(2);
                            cell.getCellStyle();
                            final CellRangeAddress range = new CellRangeAddress(StartBlockManagedEventRowNum + 2,
                                    EndBlockManagedEventRowNum + 1, 2, 4);
                            worksheet.addMergedRegion(range);
                            final Row _xRow = worksheet.getRow(StartBlockManagedEventRowNum + 2);
                            final Cell managedEventRowWithTitle = _xRow.getCell(2);
                            if (CmfConstants.MASK.equals(currentModel.getAction())) {
                                managedEventRowWithTitle.setCellValue(CmfConstants.HIDDEN_TEXT);
                            } else {
                                managedEventRowWithTitle.setCellValue(currentModel.getManagedEventName());
                            }
                            final CellStyle mergedStyle = managedEventRowWithTitle.getCellStyle();
                            mergedStyle.setAlignment(HorizontalAlignment.CENTER);
                            managedEventRowWithTitle.setCellStyle(mergedStyle);
                        }
                    }
                }
                for (final Entry<String, String> entry : nonBomValuesMap.entrySet()) {
                    final Row summaryRow = worksheet.getRow(nextRowNumber);
                    final int cellNumber = Integer.parseInt(valuesSummaryColNumber.get(entry.getKey()));
                    final Cell cell = summaryRow.getCell(cellNumber);
                    cell.getCellStyle().setFillForegroundColor(IndexedColors.WHITE.getIndex());
                    cell.setCellValue(entry.getValue());
                }
                setNextRowNumberInExcel = setNextRowNumberInExcel + 1;
            }
        }
        removeRow(worksheet, tempRowNumNewOption2);
        removeRow(worksheet, tempRowNumNewOption1);
        removeRow(worksheet, tempRowNewManagedEvent);
        removeRow(worksheet, worksheet.getLastRowNum());
        removeRow(worksheet, worksheet.getLastRowNum());
        totalAddedRows = tempRowNumNewOption1 - rowNumNewOption1;
        return totalAddedRows;
    }

    public static void removeRow(final XSSFSheet sheet, final int rowIndex) {
        final int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1, true, true);
        }
        if (rowIndex == lastRowNum) {
            final XSSFRow removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

    private static void copyRow(final XSSFWorkbook workbook, final XSSFSheet worksheet, final int sourceRowNum, final int destinationRowNum,
            final int CurOption) {
        // newRow = worksheet.getRow(destinationRowNum);
        final XSSFRow sourceRow = worksheet.getRow(sourceRowNum);
        if (destinationRowNum == worksheet.getLastRowNum()) {
            worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum() + 1, 1);
        } else {
            worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
        }
        final XSSFRow newRow = worksheet.createRow(destinationRowNum);
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            final XSSFCell oldCell = sourceRow.getCell(i);
            XSSFCell newCell = newRow.createCell(i);
            if (oldCell == null) {
                newCell = null;
                continue;
            }
            final XSSFCellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            newCell.setCellStyle(newCellStyle);
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }
            newCell.setCellType(oldCell.getCellTypeEnum());
            switch (oldCell.getCellTypeEnum()) {
            case BLANK:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case STRING:
                if (oldCell.getStringCellValue().contains("Option #")) {
                    if (CurOption == 0) {
                        newCell.setCellValue("Bom Work");
                        final CellStyle style = newCell.getCellStyle();
                        style.setWrapText(true);
                        newCell.setCellStyle(style);
                    } else
                        newCell.setCellValue("Option #" + CurOption);
                } else {
                    newCell.setCellValue(oldCell.getStringCellValue());
                }
                break;
            default:
                break;
            }
        }
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            final CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                final CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
                        cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }
        newRow.setHeight((short)700);
    }

    public void copyNewTemplateFileFromResources(final String newfileName, final String filepath) throws IOException {
        if (Desktop.isDesktopSupported()) {
            final File file = new File(System.getProperty("user.home") + "\\Downloads\\" + newfileName);
            if (file.exists())
                file.delete();
            final InputStream inputStream = ClassLoader.getSystemClassLoader().getResourceAsStream(filepath);
            // Copy file
            final OutputStream outputStream = new FileOutputStream(file);
            final byte[] buffer = new byte[1024];
            int length;
            try {
                while ((length = inputStream.read(buffer)) > 0) {
                    outputStream.write(buffer, 0, length);
                }
            } finally {
                outputStream.close();
                inputStream.close();
            }
        }
    }

    private void prepareBCMExcelTable() {
        this.bcmExcelTable = new TableView<>();
        final ArrayList<String> colList = BCMExcelConstants.BCMColumHeaders;
        final Map<String, String> colToFieldMap = BCMExcelConstants.COL_NAMES_TO_FIELDS;
        for (final String colName : colList) {
            final TableColumn<BCMModel, String> coll = new TableColumn<>(colName);
            coll.setCellValueFactory(
                    new Callback<TableColumn.CellDataFeatures<BCMModel, String>, ObservableValue<String>>() {
                        @Override
                        public ObservableValue<String> call(final CellDataFeatures<BCMModel, String> param) {
                            final SimpleStringProperty property = new SimpleStringProperty();
                            final BCMModel model = param.getValue();
                            try {
                                /*
                                 * If the User has no access to the Model Program , than hide the text.
                                 * Else set the value
                                 */
                                if (Strings.isNullOrEmpty(model.getProgramCode())) {
                                    property.set(CmfConstants.HIDDEN_TEXT);
                                } else {
                                    final Object value = PropertyUtils.getProperty(model, colToFieldMap.get(colName));
                                    if (value != null && !((String)value).trim().isEmpty()) {
                                        property.set((String)value);
                                    } else {
                                        if (!colName.contains("Feature String")) {
                                            property.set(CmfConstants.NOT_APPLICABLE);
                                        } else {
                                            property.set("VL" + param.getValue().getProdPTVL());
                                        }
                                    }
                                }
                            } catch (final Exception e) {
                                property.set("");
                            }
                            return property;
                        }
                    });
            this.bcmExcelTable.getColumns().add(coll);
        }
        this.bcmExcelTable.setItems(FXCollections.observableArrayList(this.bcmModelList));
    }
}

final class OpenTemplateFileTask implements Runnable {
    private int taskId;
    XSSFWorkbook ExcelRDFFileWorkbook;
    String XSSFWorkbook;
    FileOutputStream fileOutputStream;
    Desktop desktop;
    private String destFile;

    /**
     * Construct a OpenTemplateFileTaskClass instance
     *
     * @param workbook
     * @param destFile
     * @param fileOutputStream
     * @param desktop2
     */
    public OpenTemplateFileTask(final XSSFWorkbook workbook, final String destFile,
            final FileOutputStream fileOutputStream, final Desktop desktop2) {
        this.ExcelRDFFileWorkbook = workbook;
        this.destFile = destFile;
        this.fileOutputStream = fileOutputStream;
        this.desktop = desktop2;
    }

    @Override
    public void run() {
        try {
            this.fileOutputStream = new FileOutputStream(this.destFile);
            this.ExcelRDFFileWorkbook.write(this.fileOutputStream);
            this.fileOutputStream.flush();
            this.fileOutputStream.close();
            System.out.println("file .....");
            Desktop.getDesktop().open(new File(this.destFile));
        } catch (final IOException e) {
            BomUIUtil.ShowAlert(AlertType.ERROR, "", e.getMessage());
        }
        System.out.println("Task ID : " + this.taskId + " performed by "
                           + Thread.currentThread().getName());
    }

}
