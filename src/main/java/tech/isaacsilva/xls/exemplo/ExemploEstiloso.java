package tech.isaacsilva.xls.exemplo;

import java.awt.Font;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Workbook;

import tech.isaacsilva.xls.XlsAlignment;
import tech.isaacsilva.xls.XlsColor;
import tech.isaacsilva.xls.XlsColumn;
import tech.isaacsilva.xls.XlsFont;
import tech.isaacsilva.xls.XlsStyle;
import tech.isaacsilva.xls.XlsType;
import tech.isaacsilva.xls.XlsUtils;

public class ExemploEstiloso {

	public static void main(String[] args) throws IOException {

		List<Object[]> lista = getList();
		
		XlsFont font = new XlsFont("Arial", Font.BOLD, 11);

		// ESTILO - TITULO
		XlsStyle titleStyle = new XlsStyle();
		titleStyle.setBackgroundColor(XlsColor.BLACK);
		titleStyle.setColor(XlsColor.WHITE);
		titleStyle.setBorder(BorderStyle.THIN);
		titleStyle.setBorderColor(XlsColor.BLACK);
		titleStyle.setFont(font);

		// ESTILO - FOOTER
		XlsStyle footerStyle = titleStyle.clone();

		// ESTILO - ROW
		XlsStyle rowStyle = titleStyle.clone();
		rowStyle.setBackgroundColor(XlsColor.WHITE);
		rowStyle.setColor(XlsColor.BLACK);

		int index = 0;

		List<XlsColumn> colunas = new ArrayList<>();
		
		XlsColumn column;
		
		column = new XlsColumn("Nome", index++, XlsType.STRING);
		column.setTitleStyle(titleStyle);
		column.setRowStyle(rowStyle);
		column.setFooterEnabled(true);
		column.setFooterStyle(footerStyle);
		column.setFooterValueFn(valuesOfColumn -> "Total");
		colunas.add(column);
		
		column = new XlsColumn("CPF", index++, XlsType.CPF);
		column.setTitleStyle(titleStyle);
		column.setRowStyle(rowStyle);
		column.setFooterEnabled(true);
		column.setFooterStyle(footerStyle);
		colunas.add(column);

		column = new XlsColumn("Saldo", index, XlsType.CURRENCY);
		column.setTitleStyle(titleStyle);
		column.setRowStyle(rowStyle);
		column.getRowStyle().setAlignment(XlsAlignment.END);
		column.setFooterEnabled(true);
		column.setFooterStyle(footerStyle);
		column.getFooterStyle().setAlignment(XlsAlignment.END);
		column.setFooterValueFn(valuesOfColumn -> {

			BigDecimal total = BigDecimal.ZERO;

			for (Object value : valuesOfColumn) {
				total = total.add((BigDecimal) value);
			}

			return total;
		});
		colunas.add(column);

		Workbook workbook = XlsUtils.getWorkbook(colunas, lista);
		
		workbook.write(new FileOutputStream("excel.xlsx"));
	}

	private static List<Object[]> getList() {

		List<Object[]> lista = new ArrayList<>();

		Object[] o1 = new Object[3];

		o1[0] = "Isaac";
		o1[1] = 30065596113L;
		o1[2] = new BigDecimal("325.52");

		Object[] o2 = new Object[3];

		o2[0] = "Gabriel";
		o2[1] = 30065596113L;
		o2[2] = new BigDecimal("325.52");

		Object[] o3 = new Object[3];

		o3[0] = "Rebeca";
		o3[1] = 30065596113L;
		o3[2] = new BigDecimal("325.52");

		lista.add(o1);
		lista.add(o2);
		lista.add(o3);

		return lista;
	}

}
