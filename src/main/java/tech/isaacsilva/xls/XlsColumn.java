package tech.isaacsilva.xls;

public class XlsColumn {

	private String title;
	private int index;
	private XlsType type;

	private XlsStyle rowStyle = new XlsStyle();
	private XlsStyle titleStyle = new XlsStyle();
	private XlsStyle footerStyle = new XlsStyle();
	private boolean footerEnabled;
	private XlsValueFn footerValueFn;

	public XlsColumn(String title, int index, XlsType type) {
		this.title = title;
		this.index = index;
		this.type = type;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public XlsType getType() {
		return type;
	}

	public void setType(XlsType type) {
		this.type = type;
	}

	public XlsStyle getRowStyle() {
		return rowStyle;
	}

	public void setRowStyle(XlsStyle rowStyle) {
		this.rowStyle = rowStyle.clone();
	}

	public XlsStyle getTitleStyle() {
		return titleStyle;
	}

	public void setTitleStyle(XlsStyle titleStyle) {
		this.titleStyle = titleStyle.clone();
	}

	public XlsStyle getFooterStyle() {
		return footerStyle;
	}

	public void setFooterStyle(XlsStyle footerStyle) {
		this.footerStyle = footerStyle.clone();
	}

	public boolean isFooterEnabled() {
		return footerEnabled;
	}

	public void setFooterEnabled(boolean footerEnabled) {
		this.footerEnabled = footerEnabled;
	}

	public XlsValueFn getFooterValueFn() {
		return footerValueFn;
	}

	public void setFooterValueFn(XlsValueFn footerValueFn) {
		this.footerValueFn = footerValueFn;
	}

}
