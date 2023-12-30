package tech.isaacsilva.xls;

import java.util.List;

@FunctionalInterface
public interface XlsValueFn {

	Object getValue(List<Object> valuesOfColumn);
}
