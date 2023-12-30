package tech.isaacsilva.xls;

public enum XlsType {

	STRING(""), INT("#"), DECIMAL("#,##0.00"), DATE("DD/MM/YYYY"), PERCENT("0.00%"), CURRENCY("R$ 0.00"),
	CPF("000\".\"000\".\"000-00"), CNPJ("00\".\"000\".\"000\"/\"0000-00");

	private String format;

	XlsType(String format) {
		this.format = format;
	}

	public String getFormat() {
		return format;
	}
}
