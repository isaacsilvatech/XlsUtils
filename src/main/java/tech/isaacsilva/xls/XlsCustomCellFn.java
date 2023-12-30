package tech.isaacsilva.xls;

@FunctionalInterface
public interface XlsCustomCellFn {
	
	void apply(XlsCustomCell row);
}
