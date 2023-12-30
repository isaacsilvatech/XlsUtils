package tech.isaacsilva.xls.exemplo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

import tech.isaacsilva.xls.XlsColumn;
import tech.isaacsilva.xls.XlsType;
import tech.isaacsilva.xls.XlsUtils;

public class ExemploSimples {

	public static void main(String[] args) throws IOException {

		List<Object[]> lista = getList();

		List<XlsColumn> colunas = new ArrayList<>();

		colunas.add(new XlsColumn("Nome", 0, XlsType.STRING));
		colunas.add(new XlsColumn("CPF", 1, XlsType.CPF));
		colunas.add(new XlsColumn("Saldo", 2, XlsType.CURRENCY));

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
