package tech.isaacsilva.xls;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsUtils {

	public static byte[] toByteArray(Workbook wb) throws IOException {

		ByteArrayOutputStream outByteStream = new ByteArrayOutputStream();
		wb.write(outByteStream);

		return outByteStream.toByteArray();
	}

	public static byte[] getBytes(List<XlsColumn> columns, List<Object[]> list, XlsCustomCellFn customCellFn)
			throws IOException {

		Workbook wb = getWorkbook(columns, list, customCellFn);

		return toByteArray(wb);
	}

	private static void createFooter(Workbook wb, Sheet sheet, List<Object[]> list, List<XlsColumn> columns) {

		int lastIndex = sheet.getLastRowNum() + 1;
		Row row = sheet.createRow(lastIndex);

		for (int numberOfColumn = 0; numberOfColumn < columns.size(); numberOfColumn++) {

			XlsColumn column = columns.get(numberOfColumn);

			if (column.isFooterEnabled()) {
				Cell cell = row.createCell(numberOfColumn);
				
				XlsValueFn footerValueFn = column.getFooterValueFn();
				if(Objects.nonNull(footerValueFn)) {
					Object value = footerValueFn.getValue(getValuesOfColumn(column, list));

					setValue(cell, value, column.getType());
				}

				CellStyle style = getCellStyle(wb, column.getFooterStyle(), column.getType().getFormat());
				cell.setCellStyle(style);
			}
		}
	}

	private static List<Object> getValuesOfColumn(XlsColumn column, List<Object[]> list) {
		List<Object> valuesOfColumn = new ArrayList<>();

		for (Object[] objectRow : list) {
			valuesOfColumn.add(objectRow[column.getIndex()]);
		}

		return valuesOfColumn;
	}

	private static void createData(Workbook wb, Sheet sheet, List<Object[]> list, List<XlsColumn> columns,
			XlsCustomCellFn customCellFn, int numberOfRowInit) {

		for (int numberOfRow = 0; numberOfRow < list.size(); numberOfRow++) {

			Row row = sheet.createRow(numberOfRow + numberOfRowInit);
			Object[] objectRow = list.get(numberOfRow);

			for (int numberOfColumn = 0; numberOfColumn < columns.size(); numberOfColumn++) {

				XlsColumn column = columns.get(numberOfColumn);
				Cell cell = row.createCell(numberOfColumn);

				Object value = objectRow[column.getIndex()];

				setValue(cell, value, column.getType());

				XlsCustomCell customCell = null;
				if (Objects.nonNull(customCellFn)) {
					customCell = new XlsCustomCell();
					customCell.setObjectRow(objectRow);
					customCell.setStyle(column.getRowStyle());
					customCellFn.apply(customCell);
				}

				XlsStyle rowStyle = Objects.isNull(customCell) ? column.getRowStyle() : customCell.getStyle();
				CellStyle style = getCellStyle(wb, rowStyle, column.getType().getFormat());
				cell.setCellStyle(style);
			}
		}

	}

	private static void setValue(Cell cell, Object value, XlsType type) {

		if (value == null) {
			cell.setCellValue("");
		} else if (value instanceof String) {
			cell.setCellValue((String) value);
		} else if (value instanceof BigDecimal) {
			cell.setCellValue((((BigDecimal) value).setScale(2, RoundingMode.HALF_DOWN)).doubleValue());
		} else if (value instanceof Date) {
			cell.setCellValue((Date) value);
		} else if (value instanceof Long) {
			cell.setCellValue((Long) value);
		} else if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Short) {
			cell.setCellValue((Short) value);
		} else if (value instanceof Double) {
			cell.setCellValue((Double) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else if (value instanceof Character) {
			cell.setCellValue((Character) value);
		} else if (value instanceof Float) {
			cell.setCellValue((Float) value);
		} else if (value instanceof Byte) {
			cell.setCellValue((Byte) value);
		} else {
			cell.setCellValue(value.toString());
		}

		if (type == XlsType.PERCENT) {

			cell.setCellValue(
					((BigDecimal) value).setScale(2, RoundingMode.DOWN).divide(new BigDecimal("100")).doubleValue());
		}
	}

	private static void createTitle(Workbook wb, Sheet sheet, List<XlsColumn> columns, int index) {

		Row row = sheet.createRow(index);

		for (int i = 0; i < columns.size(); i++) {

			Cell cell = row.createCell(i);
			XlsColumn colunm = columns.get(i);

			cell.setCellValue(colunm.getTitle());
			CellStyle style = getCellStyle(wb, colunm.getTitleStyle(), XlsType.STRING.getFormat());
			cell.setCellStyle(style);
		}
	}

	private static CellStyle getCellStyle(Workbook wb, XlsStyle style, String format) {

		CellStyle cellStyle = wb.createCellStyle();
		Font font = wb.createFont();

		XlsFont xlsFont = style.getFont();

		font.setColor(style.getColor().getIndex());
		font.setFontName(xlsFont.getName());
		font.setFontHeightInPoints((short) xlsFont.getSize());

		if (xlsFont.getStyle() == XlsFont.BOLD) {
			font.setBold(true);
		} else if (xlsFont.getStyle() == XlsFont.ITALIC) {
			font.setItalic(true);
		} else if (xlsFont.getStyle() == XlsFont.BOLD_ITALIC) {
			font.setBold(true);
			font.setItalic(true);
		}

		cellStyle.setBorderTop(style.getBorder());
		cellStyle.setTopBorderColor(style.getBorderColor().getIndex());
		cellStyle.setBorderBottom(style.getBorder());
		cellStyle.setBottomBorderColor(style.getBorderColor().getIndex());
		cellStyle.setBorderLeft(style.getBorder());
		cellStyle.setLeftBorderColor(style.getBorderColor().getIndex());
		cellStyle.setBorderRight(style.getBorder());
		cellStyle.setRightBorderColor(style.getBorderColor().getIndex());

		cellStyle.setAlignment(style.getAlignment().getAlignment());

		cellStyle.setFillForegroundColor(style.getBackgroundColor().getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		cellStyle.setFont(font);

		if (format != null && !format.isEmpty()) {
			CreationHelper createHelper = wb.getCreationHelper();
			short f = createHelper.createDataFormat().getFormat(format);
			cellStyle.setDataFormat(f);
		}

		return cellStyle;
	}

	public static void autoSizeColumn(Sheet sheet, int size) {
		for (int i = 0; i < size + 1; i++) {
			sheet.autoSizeColumn(i);
		}
	}

	public static byte[] getBytes(List<XlsColumn> colunas, List<Object[]> lista) throws IOException {
		return getBytes(colunas, lista, null);
	}
	
	public static Workbook getWorkbook(List<XlsColumn> columns, List<Object[]> list) {
		return getWorkbook(columns, list, null);
	}

	public static Workbook getWorkbook(List<XlsColumn> columns, List<Object[]> list, XlsCustomCellFn customCellFn) {
		
		Workbook wb = new XSSFWorkbook();

		Sheet sheet = wb.createSheet("Planilha");

		createTitle(wb, sheet, columns, 0);
		createData(wb, sheet, list, columns, customCellFn, 1);
		createFooter(wb, sheet, list, columns);

		autoSizeColumn(sheet, columns.size());
		
		return wb;
	}
}
