package tech.isaacsilva.xls;

import org.apache.poi.ss.usermodel.BorderStyle;

public class XlsStyle implements Cloneable {

	protected XlsColor color = XlsColor.BLACK;
	protected XlsColor backgroundColor = XlsColor.WHITE;
	protected XlsColor borderColor = XlsColor.GREY_25_PERCENT;
	private BorderStyle border = BorderStyle.THIN;

	protected XlsFont font = new XlsFont("Arial", XlsFont.NORMAL, 11);
	protected XlsAlignment alignment = XlsAlignment.START;

	public XlsColor getColor() {
		return color;
	}

	public void setColor(XlsColor color) {
		this.color = color;
	}

	public XlsColor getBackgroundColor() {
		return backgroundColor;
	}

	public void setBackgroundColor(XlsColor backgroundColor) {
		this.backgroundColor = backgroundColor;
	}

	public XlsColor getBorderColor() {
		return borderColor;
	}

	public void setBorderColor(XlsColor borderColor) {
		this.borderColor = borderColor;
	}

	public XlsFont getFont() {
		return font;
	}

	public void setFont(XlsFont font) {
		this.font = font;
	}

	public XlsAlignment getAlignment() {
		return alignment;
	}

	public void setAlignment(XlsAlignment alignment) {
		this.alignment = alignment;
	}

	@Override
	public XlsStyle clone() {
		try {
			return (XlsStyle) super.clone();
		} catch (CloneNotSupportedException e) {
			throw new RuntimeException(e);
		}
	}

	public BorderStyle getBorder() {
		return border;
	}

	public void setBorder(BorderStyle border) {
		this.border = border;
	}

}
