/*
 Bandika CMS - A Java based modular Content Management System
 Copyright (C) 2009-2021 Michael Roennau

 This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 3 of the License, or (at your option) any later version.
 This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 You should have received a copy of the GNU General Public License along with this program; if not, see <http://www.gnu.org/licenses/>.
 */
package de.elbe5.file;

import com.opencsv.CSVWriter;
import de.elbe5.base.BinaryFile;
import de.elbe5.base.DateHelper;
import de.elbe5.base.LocalizedStrings;
import de.elbe5.base.StringHelper;

import java.io.IOException;
import java.io.StringWriter;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

public class CsvCreator {

    private static CsvCreator instance = null;

    public static char separator = ';';

    public static CsvCreator getInstance() {
        if (instance == null) {
            instance = new CsvCreator();
        }
        return instance;
    }

    protected String csv(String src){
        return StringHelper.toCsv(src);
    }

    protected String scsv(String src){
        return LocalizedStrings.getInstance().csv(src);
    }

    protected String csv(LocalDateTime date){
        return csv(DateHelper.toHtml(date));
    }

    protected String csv(LocalDate date){
        return csv(DateHelper.toHtmlDate(date));
    }

    public String createCSV(List<Integer> ids){
        StringWriter sw = new StringWriter();
        CSVWriter csvWriter = new CSVWriter(sw, separator, '"', '\\', "\n");
        writeContent(csvWriter, ids);
        try {
            csvWriter.close();
        }
        catch (IOException e){
            return "";
        }
        return sw.toString();
    }

    public void writeContent(CSVWriter writer, List<Integer> ids){
        writeLine(writer, "first", "second", "third");
    }

    public void writeLine(CSVWriter writer, String... strings){
        writer.writeNext(strings);
    }

    public void writeLine(CSVWriter writer, List<String> strings){
        String[] arr = new String[strings.size()];
        for (int i=0;i<strings.size();i++)
            arr[i]=strings.get(i);
        writer.writeNext(arr);
    }

    public BinaryFile getCsv(String csv, String fileName) {
        BinaryFile file=new BinaryFile();
        file.setFileName(fileName);
        file.setContentType("text/csv");
        file.setBytes(csv.getBytes(StandardCharsets.UTF_8));
        file.setFileSize(file.getBytes().length);
        return file;
    }
}
