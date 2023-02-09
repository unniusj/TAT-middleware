package io.tat.model;

import java.util.List;
import java.util.Map;
import java.util.TreeSet;

public class DualAnalyticsModel {

	private String modelName;
	private Map<String, Map<String, Integer>> dualAnalytics;
	private int total;
	private TreeSet<Integer> highlightedValues;

	public String getModelName() {
		return modelName;
	}
	public void setModelName(String modelName) {
		this.modelName = modelName;
	}
	public Map<String, Map<String, Integer>> getDualAnalytics() {
		return dualAnalytics;
	}
	public void setDualAnalytics(Map<String, Map<String, Integer>> dualAnalytics) {
		this.dualAnalytics = dualAnalytics;
	}
	public int getTotal() {
		return total;
	}
	public void setTotal(int total) {
		this.total = total;
	}
	public TreeSet<Integer> getHighlightedValues() {
		return highlightedValues;
	}
	public void setHighlightedValues(TreeSet<Integer> highlightedValues) {
		this.highlightedValues = highlightedValues;
	}
	
}
