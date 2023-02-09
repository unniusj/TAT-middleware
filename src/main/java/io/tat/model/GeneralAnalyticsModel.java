package io.tat.model;

import java.util.List;

public class GeneralAnalyticsModel {

	List<SingleAnalyticsModel> listOfSingleAnalyticsModel;
	List<DualAnalyticsModel> listOfDualAnalyticsModel;
	List<ResolutionAnalyticsModel> listOfResolutionAnalyticsModel;
	
	public List<ResolutionAnalyticsModel> getListOfResolutionAnalyticsModel() {
		return listOfResolutionAnalyticsModel;
	}
	public void setListOfResolutionAnalyticsModel(List<ResolutionAnalyticsModel> listOfResolutionAnalyticsModel) {
		this.listOfResolutionAnalyticsModel = listOfResolutionAnalyticsModel;
	}
	public List<SingleAnalyticsModel> getListOfSingleAnalyticsModel() {
		return listOfSingleAnalyticsModel;
	}
	public void setListOfSingleAnalyticsModel(List<SingleAnalyticsModel> listOfSingleAnalyticsModel) {
		this.listOfSingleAnalyticsModel = listOfSingleAnalyticsModel;
	}
	public List<DualAnalyticsModel> getListOfDualAnalyticsModel() {
		return listOfDualAnalyticsModel;
	}
	public void setListOfDualAnalyticsModel(List<DualAnalyticsModel> listOfDualAnalyticsModel) {
		this.listOfDualAnalyticsModel = listOfDualAnalyticsModel;
	}
	
}
