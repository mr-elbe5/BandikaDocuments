/*
 Bandika CMS - A Java based modular Content Management System
 Copyright (C) 2009-2021 Michael Roennau

 This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 3 of the License, or (at your option) any later version.
 This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 You should have received a copy of the GNU General Public License along with this program; if not, see <http://www.gnu.org/licenses/>.
 */
package de.elbe5.file;

import de.elbe5.base.Log;
import de.elbe5.base.StringHelper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

public class XlsxCreator {

    Workbook workbook = new XSSFWorkbook();
    CellStyle headerCellStyle = workbook.createCellStyle();
    CellStyle defaultCellStyle = workbook.createCellStyle();

    protected Sheet currentSheet;
    protected int currentRowCount;
    protected int currentRow;

    protected String xml(String src) {
        return StringHelper.toXml(src);
    }

    public XlsxCreator(){
        workbook = new XSSFWorkbook();
        headerCellStyle = workbook.createCellStyle();
        setHeaderStyle();
        defaultCellStyle = workbook.createCellStyle();
        setDefaultStyle();
    }

    public byte[] createExcelFile() {
        addSheets();
        try{
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        } catch (IOException e) {
            Log.error("could not create xlsx output", e);
            return null;
        }
    }

    protected void setHeaderStyle(){
        headerCellStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setBorderTop(BorderStyle.THIN);
        headerCellStyle.setTopBorderColor(IndexedColors.BLACK.index);
        headerCellStyle.setBorderRight(BorderStyle.THIN);
        headerCellStyle.setRightBorderColor(IndexedColors.BLACK.index);
        headerCellStyle.setBorderBottom(BorderStyle.THIN);
        headerCellStyle.setBottomBorderColor(IndexedColors.BLACK.index);
        headerCellStyle.setBorderLeft(BorderStyle.THIN);
        headerCellStyle.setLeftBorderColor(IndexedColors.BLACK.index);
    }

    protected void setDefaultStyle(){
        defaultCellStyle.setBorderTop(BorderStyle.THIN);
        defaultCellStyle.setTopBorderColor(IndexedColors.BLACK.index);
        defaultCellStyle.setBorderRight(BorderStyle.THIN);
        defaultCellStyle.setRightBorderColor(IndexedColors.BLACK.index);
        defaultCellStyle.setBorderBottom(BorderStyle.THIN);
        defaultCellStyle.setBottomBorderColor(IndexedColors.BLACK.index);
        defaultCellStyle.setBorderLeft(BorderStyle.THIN);
        defaultCellStyle.setLeftBorderColor(IndexedColors.BLACK.index);
    }

    protected void addSheets(){

    }

    protected void addSheet(String name, String... headers){
        currentSheet = workbook.createSheet();
        currentRowCount = headers.length;
        Row headerRow = currentSheet.createRow(0);
        currentRow = 0;
        for(int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(StringHelper.toXml(headers[i]));
            cell.setCellStyle(headerCellStyle);
        }
    }

    protected void autoSizeSheet(){
        for (int i=0; i< currentRowCount; i++){
            currentSheet.autoSizeColumn(i);
        }
    }

    protected void addRow(String... values) {
        currentRow++;
        Row row = currentSheet.createRow(currentRow);
        if (values.length != currentRowCount){
            Log.warn("row length mismatch in xlsx");
        }
        for(int i = 0; i < values.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(StringHelper.toXml(values[i]));
            cell.setCellStyle(defaultCellStyle);
        }
    }

}
