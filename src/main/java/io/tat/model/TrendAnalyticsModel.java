package io.tat.model;

import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

public class TrendAnalyticsModel {

	List<TrendDateModel> trendDateModelList;
	List<TrendDimModel> trendDimModelList;
	TreeSet<String> years;
	TreeSet<String> dims;
	Map<String, Set<String>> allDims;
	
	public Map<String, Set<String>> getAllDims() {
		return allDims;
	}
	public void setAllDims(Map<String, Set<String>> allDims) {
		this.allDims = allDims;
	}
	public List<TrendDateModel> getTrendDateModelList() {
		return trendDateModelList;
	}
	public void setTrendDateModelList(List<TrendDateModel> trendDateModelList) {
		this.trendDateModelList = trendDateModelList;
	}
	public List<TrendDimModel> getTrendDimModelList() {
		return trendDimModelList;
	}
	public void setTrendDimModelList(List<TrendDimModel> trendDimModelList) {
		this.trendDimModelList = trendDimModelList;
	}
	public TreeSet<String> getYears() {
		return years;
	}
	public void setYears(TreeSet<String> years) {
		this.years = years;
	}
	public TreeSet<String> getDims() {
		return dims;
	}
	public void setDims(TreeSet<String> dims) {
		this.dims = dims;
	}
	
}
