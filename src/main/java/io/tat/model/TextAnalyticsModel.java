package io.tat.model;

public class TextAnalyticsModel {

    private String listName;

    private String headerId;
    private String listId;

    private String[] headerData;
    private String[] data;

    private int headerTotal;
    private int childTotal;
    private int headerPer;
    private int childPer;

    public TextAnalyticsModel() {
    	System.out.println("hello");
    }

	public TextAnalyticsModel(String listName, String headerId, String listId, String[] headerData, String[] data,
			int headerTotal, int childTotal, int headerPer, int childPer) {
		super();
		this.listName = listName;
		this.headerId = headerId;
		this.listId = listId;
		this.headerData = headerData;
		this.data = data;
		this.headerTotal = headerTotal;
		this.childTotal = childTotal;
		this.headerPer = headerPer;
		this.childPer = childPer;
	}


	public String getListName() {
		return listName;
	}

	public void setListName(String listName) {
		this.listName = listName;
	}

	public String getHeaderId() {
		return headerId;
	}

	public void setHeaderId(String headerId) {
		this.headerId = headerId;
	}

	public String getListId() {
		return listId;
	}

	public void setListId(String listId) {
		this.listId = listId;
	}

	public String[] getHeaderData() {
		return headerData;
	}

	public void setHeaderData(String[] headerData) {
		this.headerData = headerData;
	}

	public String[] getData() {
		return data;
	}

	public void setData(String[] data) {
		this.data = data;
	}

	public int getHeaderTotal() {
		return headerTotal;
	}

	public void setHeaderTotal(int headerTotal) {
		this.headerTotal = headerTotal;
	}

	public int getChildTotal() {
		return childTotal;
	}

	public void setChildTotal(int childTotal) {
		this.childTotal = childTotal;
	}

	public int getHeaderPer() {
		return headerPer;
	}

	public void setHeaderPer(int headerPer) {
		this.headerPer = headerPer;
	}

	public int getChildPer() {
		return childPer;
	}

	public void setChildPer(int childPer) {
		this.childPer = childPer;
	}
    
    
}
