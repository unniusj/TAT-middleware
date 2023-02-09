package io.tat.model;

import java.util.Map;
import java.util.TreeSet;

public class ResolutionAnalyticsModel {

	private String modelName;
	private Map<String, String> analytics;
	private double total;
	private TreeSet<String> highlightedValues;

	public String getModelName() {
		return modelName;
	}
	public void setModelName(String modelName) {
		this.modelName = modelName;
	}
	public Map<String, String> getAnalytics() {
		return analytics;
	}
	public void setAnalytics(Map<String, String> analytics) {
		this.analytics = analytics;
	}
	public double getTotal() {
		return total;
	}
	public void setTotal(double total) {
		this.total = total;
	}
	public TreeSet<String> getHighlightedValues() {
		return highlightedValues;
	}
	public void setHighlightedValues(TreeSet<String> highlightedValues) {
		this.highlightedValues = highlightedValues;
	}
	
}
