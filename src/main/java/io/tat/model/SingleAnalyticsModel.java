package io.tat.model;

import java.util.Map;
import java.util.TreeSet;

public class SingleAnalyticsModel {

	private String modelName;
	private Map<String, Integer> analytics;
	private int total;
	private TreeSet<Integer> highlightedValues;

	public String getModelName() {
		return modelName;
	}
	public void setModelName(String modelName) {
		this.modelName = modelName;
	}
	public Map<String, Integer> getAnalytics() {
		return analytics;
	}
	public void setAnalytics(Map<String, Integer> analytics) {
		this.analytics = analytics;
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
