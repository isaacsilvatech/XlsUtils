# XlsUtils - Gere Excels com Facilidade

## Visão geral
Uma biblioteca Java leve para simplificar a geração de planilhas Excel com estilos personalizáveis. Esta biblioteca fornece uma interface fácil de usar para criar arquivos Excel a partir de uma lista de dados, permitindo que você se concentre no conteúdo da sua planilha em vez de lidar com as complexidades da criação de arquivos Excel.

## Características
- **API simples:** API minimalista e fácil de entender para geração de arquivos Excel.
- **Estilos personalizáveis:** defina estilos para títulos, linhas e rodapés para melhorar a aparência visual da sua planilha.
- **Configuração flexível de colunas:** configure cada coluna com estilos, tipos e cálculos de rodapé específicos.

## Exemplo simples:
![Capturar](https://github.com/isaacsilvatech/XlsUtils/assets/145171555/2c952349-9792-49c0-a784-d16cb7e119c4)
```
		List<Object[]> lista = getList();
		
		List<XlsColumn> colunas = new ArrayList<>();
		
		colunas.add(new XlsColumn("Nome", 0, XlsType.STRING));
		colunas.add(new XlsColumn("CPF", 1, XlsType.CPF));
		colunas.add(new XlsColumn("Saldo", 2, XlsType.CURRENCY));
		
		Workbook workbook = XlsUtils.getWorkbook(colunas, lista);
				
		workbook.write(new FileOutputStream("excel.xlsx"));
```

## Exemplo estiloso:
![Capturar2](https://github.com/isaacsilvatech/XlsUtils/assets/145171555/ab8ef175-fec5-43f8-962f-56db4fe77418)
```
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
```

