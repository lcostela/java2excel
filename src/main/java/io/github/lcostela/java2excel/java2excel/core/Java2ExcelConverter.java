/*
 * MIT License
 *
 * Copyright (c) 2018 Luis Costela
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 */

package io.github.lcostela.java2excel.java2excel.core;

import io.github.lcostela.java2excel.java2excel.annotatios.CellHeader;
import io.github.lcostela.java2excel.java2excel.annotatios.CellOrder;
import io.github.lcostela.java2excel.java2excel.exception.IllegalTypeOfAttribute;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Created by Luis Costela.
 */
public class Java2ExcelConverter<T> {

    private Class<T> clazz;
    private final Logger logger = Logger.getLogger(Java2ExcelConverter.class.getName());

    public Java2ExcelConverter(Class<T> clazz) {
        this.clazz = clazz;
    }

    private Class<T> obtainClass() {
        return clazz;
    }

    public Workbook toWorkBook(List<T> itemReport, String name) throws IllegalTypeOfAttribute {
        checkClassEligibleToConvert(obtainClass());
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(name);
        Row row = sheet.createRow(0);
        generateHeaders(obtainClass(), row);
        int i = 1;
        for (T item : itemReport) {
            row = sheet.createRow(i);
            generateRow(item, row);
            i++;
        }
        final int numberOfColumns = getNumberOfColumns(obtainClass());
        autoSizeColumns(sheet, numberOfColumns);
        sheet.setAutoFilter(new CellRangeAddress(0, i, 0, numberOfColumns));
        return workbook;
    }

    private void checkClassEligibleToConvert(Class<T> clazz) throws IllegalTypeOfAttribute {
        List<Field> attributes = FieldUtils.getAllFieldsList(clazz);
        final Class[] classes = {Long.class, String.class, Double.class, LocalDate.class};
        try {
            compareClasses(attributes, classes);
        } catch (IllegalTypeOfAttribute illegalTypeOfAttribute) {
            //todo message
            logger.log(Level.SEVERE, "Message...", illegalTypeOfAttribute);
            throw illegalTypeOfAttribute;
        }
    }

    private void compareClasses(List<Field> attributes, Class[] classes) throws IllegalTypeOfAttribute {
        for (Field attribute : attributes) {
            boolean cast = false;
            for (Class klazz : classes) {
                if (fieldIsFromClass(attribute, klazz)) {
                    cast = true;
                    break;
                }
            }
            //todo: revisar
            if (!cast && !attribute.getType().isEnum()) {
                throw new IllegalTypeOfAttribute();
            }
        }
    }

    private boolean fieldIsFromClass(Field field, Class clazz) {
        return field.getType().equals(clazz);
    }

    private int getNumberOfColumns(Class<T> item) {

        int numberOfColumns = 0;
        Method[] m = new Method[0];
        m = item.getDeclaredMethods();
        for (Method method : m) {
            if (method.getAnnotation(CellOrder.class) != null) {
                numberOfColumns++;
            }
        }
        return numberOfColumns;
    }

    private void generateRow(T item, Row row) {
        try {
            Class c = item.getClass();
            Method[] m = c.getDeclaredMethods();
            for (Method method : m) {
                if (method.getAnnotation(CellOrder.class) != null) {
                    int columnNumber = method.getAnnotation(CellOrder.class).value();
                    Cell cel = row.createCell(columnNumber);
                    Object value = method.invoke(item);
                    if (value != null) {
                        if (value instanceof String) {
                            cel.setCellValue((String) value);
                        } else if (value instanceof Long) {
                            cel.setCellValue((Long) value);
                        } else if (value instanceof Integer) {
                            cel.setCellValue((Integer) value);
                        } else if (value instanceof Double) {
                            cel.setCellValue((Double) value);
                        } else if (value instanceof LocalDate) {
                            cel.setCellValue(value.toString());
                        } else if (value instanceof Enum<?>) {
                            cel.setCellValue(((Enum) value).name());
                        }
                    }
                }
            }
        } catch (IllegalAccessException e) {
            //todo message
            logger.log(Level.SEVERE, "Message...", e);
        } catch (InvocationTargetException e) {
            //todo message
            logger.log(Level.SEVERE, "Message2...", e);
        }
    }

    private void generateHeaders(Class<T> item, Row row) {
        Method[] m = item.getDeclaredMethods();
        for (Method method : m) {
            if (method.getAnnotation(CellOrder.class) != null) {
                int columnNumber = method.getAnnotation(CellOrder.class).value();
                row.createCell(columnNumber).setCellValue(method.getAnnotation(CellHeader.class).value());
            }
        }
    }

    private void autoSizeColumns(Sheet sheet, int numberOfColumns) {
        for (int i = 0; i < numberOfColumns; i++) {
            sheet.autoSizeColumn(i);
        }
    }
}
