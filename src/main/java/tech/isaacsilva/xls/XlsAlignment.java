package tech.isaacsilva.xls;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

public enum XlsAlignment {

	START(HorizontalAlignment.LEFT), 
	CENTER(HorizontalAlignment.CENTER), 
	END(HorizontalAlignment.RIGHT);

	private HorizontalAlignment alignment;

	XlsAlignment(HorizontalAlignment alignment) {
		this.alignment = alignment;
	}

	public HorizontalAlignment getAlignment() {
		return alignment;
	}

}
