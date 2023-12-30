# XlsUtils

## Visão geral
Uma biblioteca Java leve para simplificar a geração de planilhas Excel com estilos personalizáveis. Esta biblioteca fornece uma interface fácil de usar para criar arquivos Excel a partir de uma lista de dados, permitindo que você se concentre no conteúdo da sua planilha em vez de lidar com as complexidades da criação de arquivos Excel.

## Características
- **API simples:** API minimalista e fácil de entender para geração de arquivos Excel.
- **Estilos personalizáveis:** defina estilos para títulos, linhas e rodapés para melhorar a aparência visual da sua planilha.
- **Configuração flexível de colunas:** configure cada coluna com estilos, tipos e cálculos de rodapé específicos.

## Excemplo simples:
'''
List<Object[]> lista = getList();

List<XlsColumn> colunas = new ArrayList<>();

colunas.add(new XlsColumn("Nome", 0, XlsType.STRING));
colunas.add(new XlsColumn("CPF", 1, XlsType.CPF));
colunas.add(new XlsColumn("Saldo", 2, XlsType.CURRENCY));

Workbook workbook = XlsUtils.getWorkbook(colunas, lista);
		
workbook.write(new FileOutputStream("excel.xlsx"));
'''

![Capturar](https://github.com/isaacsilvatech/XlsUtils/assets/145171555/87525e77-dcc0-421f-bcb8-9b4b0c22799a)
