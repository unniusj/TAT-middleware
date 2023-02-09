package io.tat.model;

import java.util.Map;

public class ColumnModel {

	Map<String, Integer> columnNames;
	Map<String, String> columnTypes;
	
	public Map<String, Integer> getColumnNames() {
		return columnNames;
	}
	public void setColumnNames(Map<String, Integer> columnNames) {
		this.columnNames = columnNames;
	}
	public Map<String, String> getColumnTypes() {
		return columnTypes;
	}
	public void setColumnTypes(Map<String, String> columnTypes) {
		this.columnTypes = columnTypes;
	}
	
}
