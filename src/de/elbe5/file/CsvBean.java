/*
 Bandika CMS - A Java based modular Content Management System
 Copyright (C) 2009-2021 Michael Roennau

 This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 3 of the License, or (at your option) any later version.
 This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 You should have received a copy of the GNU General Public License along with this program; if not, see <http://www.gnu.org/licenses/>.
 */
package de.elbe5.file;

import com.opencsv.CSVWriter;
import de.elbe5.base.StringHelper;
import de.elbe5.database.DbBean;

import java.io.IOException;
import java.io.StringWriter;

public class CsvBean extends DbBean {

    private static CsvBean instance = null;

    public static CsvBean getInstance() {
        if (instance == null) {
            instance = new CsvBean();
        }
        return instance;
    }

    protected String xml(String src){
        return StringHelper.toXml(src);
    }

    public String createCSV(){
        StringWriter sw = new StringWriter();
        CSVWriter csvWriter = new CSVWriter(sw);
        writeContent(csvWriter);
        try {
            csvWriter.close();
        }
        catch (IOException e){
            return "";
        }
        return sw.toString();
    }

    public void writeContent(CSVWriter writer){
        writeLine(writer, "first", "second", "third");
    }

    public void writeLine(CSVWriter writer, String... strings){
        writer.writeNext(strings);
    }

}
