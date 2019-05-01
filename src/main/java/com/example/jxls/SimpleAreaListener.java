package com.example.jxls;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.jxls.area.XlsArea;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;

public class SimpleAreaListener implements AreaListener {

    private XlsArea area;
    private PoiTransformer transformer;
    private final CellRef paymentCell = new CellRef("Hárok1!B2");

    public SimpleAreaListener(Transformer transformer) {
        this.transformer = (PoiTransformer) transformer;
    }


    @Override
    public void beforeApplyAtCell(CellRef cellRef, Context context) {

    }

    @Override
    public void beforeTransformCell(CellRef cellRef, CellRef cellRef1, Context context) {

    }

    @Override
    public void afterApplyAtCell(CellRef cellRef, Context context) {

    }

    @Override
    public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
        System.out.println("Source: " + srcCell.getCellName() + "Target: " + targetCell.getCellName());
        if(paymentCell.equals(srcCell)){
            Price price = (Price) context.getVar("c");
            if( price.getTypeOfValue() == TypeOfValue.HIGH ){
                highlightCell(targetCell,HSSFColor.HSSFColorPredefined.RED.getIndex());
            }

            if( price.getTypeOfValue() == TypeOfValue.LOW ){
                highlightCell(targetCell,HSSFColor.HSSFColorPredefined.GREEN.getIndex());
            }
        }
    }
    private void highlightCell(CellRef cellRef, short color) {
        Workbook workbook = transformer.getWorkbook();
        Sheet sheet = workbook.getSheet(cellRef.getSheetName());
        Cell cell = sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
        CellStyle oldstyle = cell.getCellStyle();
        CellStyle newCellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setColor(color);
        newCellStyle.setFont(font);
        newCellStyle.setBorderBottom(oldstyle.getBorderBottom());
        newCellStyle.setBorderLeft(oldstyle.getBorderLeft());
        newCellStyle.setBorderRight(oldstyle.getBorderRight());
        newCellStyle.setBorderTop(oldstyle.getBorderTop());
        newCellStyle.setAlignment(oldstyle.getAlignment());
        newCellStyle.setAlignment(oldstyle.getAlignment());
        cell.setCellStyle(newCellStyle);
        cell.setCellValue("●");

    }
}
