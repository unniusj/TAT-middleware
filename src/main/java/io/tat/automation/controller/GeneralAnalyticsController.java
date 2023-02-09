package io.tat.automation.controller;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DateFormatSymbols;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.stream.Collectors;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.functions.Offset;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.fasterxml.jackson.databind.ObjectMapper;

import io.tat.model.ColumnModel;
import io.tat.model.DualAnalyticsModel;
import io.tat.model.GeneralAnalyticsModel;
import io.tat.model.ResolutionAnalyticsModel;
import io.tat.model.SingleAnalyticsModel;
import io.tat.model.TextAnalyticsModel;
import io.tat.model.TrendAnalyticsModel;
import io.tat.model.TrendDateModel;
import io.tat.model.TrendDimModel;

@RestController
@CrossOrigin(origins = "http://localhost:4200")
public class GeneralAnalyticsController {

	private Sheet datatypeSheet = null;
	Map<String, Integer> columnNames = new TreeMap<>();
	Map<String, String> columnTypes = new HashMap<>();
	Map<String, String> dataTypeColumns = new HashMap<>();
	Map<String, Set<String>> allDims = new HashMap<>();
	private String DATE = "Date";
	private String STRING = "String";
	private String NUMERIC = "Numeric";
	private String BOOLEAN = "Boolean";
	private DecimalFormat df2 = new DecimalFormat("#.##");
	ArrayList<String> choosenColumns = null;
	GeneralAnalyticsModel generalAnalyticsModel = new GeneralAnalyticsModel();
	int dateIndex;
	int dimIndex;
	int startDateIndex = -1;
	int endDateIndex = -1;

	@RequestMapping(method = RequestMethod.GET, value = "/check")
	public String checkCheck() {
		return "It's up!";
	}

	@RequestMapping(method = RequestMethod.POST, value = "/submitexcel")
	public ColumnModel getFile(@RequestParam MultipartFile file) throws InvalidFormatException {
		Map<String, Integer> columnNames = new HashMap<>();

		columnNames.clear();
		columnTypes.clear();
		dataTypeColumns.clear();
		datatypeSheet = null;
		try {

			FileInputStream excelFile = (FileInputStream) file.getInputStream();
			Workbook workbook = new XSSFWorkbook(excelFile);
			datatypeSheet = workbook.getSheetAt(0);

			columnTypes = findDataTypes();
			columnNames = getColumnNames();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println(file.getOriginalFilename() + " Uploaded");
		ColumnModel columnModel = new ColumnModel();
		columnModel.setColumnNames(columnNames);
		columnModel.setColumnTypes(columnTypes);
		return columnModel;
	}

	@RequestMapping(method = RequestMethod.GET, value = "/findDuplicates")
	public String checkForDuplicates(@RequestParam String pkColumn) {
		String result = "No Duplicates";

		List<String> pkList = new ArrayList<>();

		Iterator<Row> iterator = datatypeSheet.iterator();
		iterator.next();
		int count = 0;
		while (iterator.hasNext()) {
			System.out.println("current line : "+count++);
			Row currentRow = iterator.next();
			String pkValue = currentRow.getCell(Integer.valueOf(pkColumn)).getStringCellValue();
			if (pkList.contains(pkValue)) {
				result = "Duplicates found. Please correct and import the file again.";
				return result;
			} else
				pkList.add(pkValue);
		}
		return result;
	}

	private Map<String, Integer> getColumnNames() {
		allDims.clear();
		Iterator<Row> iterator = datatypeSheet.iterator();

		Row currentRow = iterator.next();
		Iterator<Cell> cellIterator = currentRow.iterator();
		int colNum = 0;
		columnNames.clear();
		Map<Integer, String> cols = new HashMap<>();
		while (cellIterator.hasNext()) {

			Cell currentCell = cellIterator.next();
			columnNames.put(currentCell.getStringCellValue(), colNum);
			cols.put(colNum, currentCell.getStringCellValue());
			colNum++;

		}

		while (iterator.hasNext()) {
			int colNo = 0;
			Row row = iterator.next();
			Iterator<Cell> cellIter = row.iterator();
			while (cellIter.hasNext()) {
				Cell currentCell = cellIter.next();
				try {
					if (getDatatypeOfColumn(cols.get(colNo)).equals("String")) {
						if (allDims.containsKey(cols.get(colNo))) {
							Set<String> values = allDims.get(cols.get(colNo));
							values.add(currentCell.getStringCellValue());
							allDims.put(cols.get(colNo), values);
						} else {
							Set<String> values = new HashSet<>();
							values.add(currentCell.getStringCellValue());
							allDims.put(cols.get(colNo), values);
						}
					}
					colNo++;
				} catch (Exception e) {
					break;
				}
			}

		}
		return columnNames;
	}

	private Map<String, String> findDataTypes() {

		Iterator<Row> iterator = datatypeSheet.iterator();

		Row headerRow = iterator.next();
		Row firstRow = iterator.next();
		Iterator<Cell> cellIterator = firstRow.iterator();
		int colNum = 1;
		dataTypeColumns.clear();
		while (cellIterator.hasNext()) {
			String type = STRING;
			Cell currentCell = cellIterator.next();
			System.out.println(headerRow.getCell(colNum - 1).getStringCellValue());
			try {
				if (DateUtil.isCellDateFormatted(currentCell))
					type = DATE;
				else if (currentCell.getCellType() == CellType.NUMERIC) {
					type = NUMERIC;
				} else if (currentCell.getCellType() == CellType.BOOLEAN) {
					type = BOOLEAN;
				} else {
					type = STRING;
				}
			} catch (Exception e) {
				if (currentCell.getCellType() == CellType.NUMERIC) {
					type = NUMERIC;
				} else if (currentCell.getCellType() == CellType.BOOLEAN) {
					type = BOOLEAN;
				} else {
					type = STRING;
				}
			}

			dataTypeColumns.put(headerRow.getCell(colNum - 1).getStringCellValue(), type);
			colNum++;
		}
//		System.out.println(dataTypeColumns);
		return dataTypeColumns;
	}

	private DualAnalyticsModel doDualAnalytics(List<String> selectedColumns) {

		DualAnalyticsModel dualAnalyticsModel = new DualAnalyticsModel();
		List<Integer> selColIndexes = new ArrayList<Integer>();
		selectedColumns.forEach(k -> selColIndexes.add(columnNames.get(k)));

		Map<String, Map<String, Integer>> dualAnalytics = new HashMap<>();

		Iterator<Row> iterator = datatypeSheet.iterator();
		iterator.next();
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();
			Cell cell = currentRow.getCell(selColIndexes.get(0));
			Cell cell1 = currentRow.getCell(selColIndexes.get(1));
			String datatype1 = getDatatypeOfColumn(selectedColumns.get(0));
			String datatype2 = getDatatypeOfColumn(selectedColumns.get(1));

			if (null != cell && null != cell1) {
				switch (datatype1) {
				case "String":
					if (dualAnalytics.containsKey(cell.getStringCellValue())) {
						Map<String, Integer> innerMap = dualAnalytics.get(cell.getStringCellValue());
						switch (datatype2) {
						case "String":
							if (innerMap.containsKey(cell1.getStringCellValue())) {
								innerMap.put(cell1.getStringCellValue(), innerMap.get(cell1.getStringCellValue()) + 1);
							} else {
								innerMap.put(cell1.getStringCellValue(), 1);
							}
							break;
						case "Numeric":
							if (innerMap.containsKey(String.valueOf(cell1.getNumericCellValue()))) {
								innerMap.put(String.valueOf(cell1.getNumericCellValue()),
										innerMap.get(String.valueOf(cell1.getNumericCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getNumericCellValue()), 1);
							}
							break;
						case "Boolean":
							if (innerMap.containsKey(String.valueOf(cell1.getBooleanCellValue()))) {
								innerMap.put(String.valueOf(cell1.getBooleanCellValue()),
										innerMap.get(String.valueOf(cell1.getBooleanCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getBooleanCellValue()), 1);
							}
							break;
						case "Date":
							if (innerMap.containsKey(String.valueOf(cell1.getDateCellValue()))) {
								innerMap.put(String.valueOf(cell1.getDateCellValue()),
										innerMap.get(String.valueOf(cell1.getDateCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getDateCellValue()), 1);
							}
							break;
						default:
							break;
						}

						innerMap = sortMapByValue(innerMap);
						dualAnalytics.put(cell.getStringCellValue(), innerMap);
					} else {
						dualAnalytics.put(cell.getStringCellValue(), new HashMap<String, Integer>());
					}
					break;
				case "Numeric":
					if (dualAnalytics.containsKey(String.valueOf(cell.getNumericCellValue()))) {
						Map<String, Integer> innerMap = dualAnalytics.get(String.valueOf(cell.getNumericCellValue()));
						switch (datatype2) {
						case "String":
							if (innerMap.containsKey(cell1.getStringCellValue())) {
								innerMap.put(cell1.getStringCellValue(), innerMap.get(cell1.getStringCellValue()) + 1);
							} else {
								innerMap.put(cell1.getStringCellValue(), 1);
							}
							break;
						case "Numeric":
							if (innerMap.containsKey(String.valueOf(cell1.getNumericCellValue()))) {
								innerMap.put(String.valueOf(cell1.getNumericCellValue()),
										innerMap.get(String.valueOf(cell1.getNumericCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getNumericCellValue()), 1);
							}
							break;
						case "Boolean":
							if (innerMap.containsKey(String.valueOf(cell1.getBooleanCellValue()))) {
								innerMap.put(String.valueOf(cell1.getBooleanCellValue()),
										innerMap.get(String.valueOf(cell1.getBooleanCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getBooleanCellValue()), 1);
							}
							break;
						case "Date":
							if (innerMap.containsKey(String.valueOf(cell1.getDateCellValue()))) {
								innerMap.put(String.valueOf(cell1.getDateCellValue()),
										innerMap.get(String.valueOf(cell1.getDateCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getDateCellValue()), 1);
							}
							break;
						default:
							break;
						}

						innerMap = sortMapByValue(innerMap);
						dualAnalytics.put(String.valueOf(cell.getNumericCellValue()), innerMap);
					} else {
						dualAnalytics.put(String.valueOf(cell.getNumericCellValue()), new HashMap<String, Integer>());
					}
					break;
				case "Boolean":
					if (dualAnalytics.containsKey(String.valueOf(cell.getBooleanCellValue()))) {
						Map<String, Integer> innerMap = dualAnalytics.get(String.valueOf(cell.getBooleanCellValue()));
						switch (datatype2) {
						case "String":
							if (innerMap.containsKey(cell1.getStringCellValue())) {
								innerMap.put(cell1.getStringCellValue(), innerMap.get(cell1.getStringCellValue()) + 1);
							} else {
								innerMap.put(cell1.getStringCellValue(), 1);
							}
							break;
						case "Numeric":
							if (innerMap.containsKey(String.valueOf(cell1.getNumericCellValue()))) {
								innerMap.put(String.valueOf(cell1.getNumericCellValue()),
										innerMap.get(String.valueOf(cell1.getNumericCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getNumericCellValue()), 1);
							}
							break;
						case "Boolean":
							if (innerMap.containsKey(String.valueOf(cell1.getBooleanCellValue()))) {
								innerMap.put(String.valueOf(cell1.getBooleanCellValue()),
										innerMap.get(String.valueOf(cell1.getBooleanCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getBooleanCellValue()), 1);
							}
							break;
						case "Date":
							if (innerMap.containsKey(String.valueOf(cell1.getDateCellValue()))) {
								innerMap.put(String.valueOf(cell1.getDateCellValue()),
										innerMap.get(String.valueOf(cell1.getDateCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getDateCellValue()), 1);
							}
							break;
						default:
							break;
						}

						innerMap = sortMapByValue(innerMap);
						dualAnalytics.put(String.valueOf(cell.getBooleanCellValue()), innerMap);
					} else {
						dualAnalytics.put(String.valueOf(cell.getBooleanCellValue()), new HashMap<String, Integer>());
					}
					break;
				case "Date":
					if (dualAnalytics.containsKey(String.valueOf(cell.getDateCellValue()))) {
						Map<String, Integer> innerMap = dualAnalytics.get(String.valueOf(cell.getDateCellValue()));
						switch (datatype2) {
						case "String":
							if (innerMap.containsKey(cell1.getStringCellValue())) {
								innerMap.put(cell1.getStringCellValue(), innerMap.get(cell1.getStringCellValue()) + 1);
							} else {
								innerMap.put(cell1.getStringCellValue(), 1);
							}
							break;
						case "Numeric":
							if (innerMap.containsKey(String.valueOf(cell1.getNumericCellValue()))) {
								innerMap.put(String.valueOf(cell1.getNumericCellValue()),
										innerMap.get(String.valueOf(cell1.getNumericCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getNumericCellValue()), 1);
							}
							break;
						case "Boolean":
							if (innerMap.containsKey(String.valueOf(cell1.getBooleanCellValue()))) {
								innerMap.put(String.valueOf(cell1.getBooleanCellValue()),
										innerMap.get(String.valueOf(cell1.getBooleanCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getBooleanCellValue()), 1);
							}
							break;
						case "Date":
							if (innerMap.containsKey(String.valueOf(cell1.getDateCellValue()))) {
								innerMap.put(String.valueOf(cell1.getDateCellValue()),
										innerMap.get(String.valueOf(cell1.getDateCellValue())) + 1);
							} else {
								innerMap.put(String.valueOf(cell1.getDateCellValue()), 1);
							}
							break;
						default:
							break;
						}

						innerMap = sortMapByValue(innerMap);
						dualAnalytics.put(String.valueOf(cell.getDateCellValue()), innerMap);
					} else {
						dualAnalytics.put(String.valueOf(cell.getDateCellValue()), new HashMap<String, Integer>());
					}
					break;
				default:
					break;
				}
			} else {
				System.out.println("Wrong Cell");
			}

		}

		Map<String, Map<String, Integer>> sortedDAnalysis = new LinkedHashMap<String, Map<String, Integer>>();

		Map<String, Integer> helperMap = new LinkedHashMap<String, Integer>();

		Map<String, Integer> helperMapTemp = new LinkedHashMap<String, Integer>();

		TreeSet<Integer> allValues = new TreeSet<>();
		int dAnalyticsTotal = 0;

		for (Map.Entry<String, Map<String, Integer>> entry : dualAnalytics.entrySet()) {
			int total = 0;
			for (Map.Entry<String, Integer> innerEntry : entry.getValue().entrySet()) {
				total = total + innerEntry.getValue();
				allValues.add(innerEntry.getValue());
			}
			helperMapTemp.put(entry.getKey(), total);
			dAnalyticsTotal = dAnalyticsTotal + total;
		}
//		allValues = (TreeSet<Integer>)allValues.descendingSet();

		TreeSet<Integer> highlightedValues = new TreeSet<>();
		int i = 1;
		while (allValues.size() > 0 && i <= 3) {
			i++;
			highlightedValues.add(allValues.pollLast());
		}

		helperMap = sortMapByValue(helperMapTemp);
		helperMap.forEach((m, n) -> {
			sortedDAnalysis.put(m, dualAnalytics.get(m));
		});
//		sortedDAnalysis.put("Grand Total", value);
		dualAnalyticsModel.setModelName(selectedColumns.get(0) + "-" + selectedColumns.get(1));
		dualAnalyticsModel.setDualAnalytics(sortedDAnalysis);
		dualAnalyticsModel.setTotal(dAnalyticsTotal);
		dualAnalyticsModel.setHighlightedValues(highlightedValues);
		return dualAnalyticsModel;
	}

	@SuppressWarnings("unchecked")
	@RequestMapping(method = RequestMethod.POST, path = "/postMultipleColumns")
	public void postMSelectedColumns(String choosenNames, String startDateColumn, String endDateColumn) {

		ObjectMapper mapper = new ObjectMapper();
		try {
			choosenColumns = mapper.readValue(choosenNames, ArrayList.class);
		} catch (IOException e) {
			e.printStackTrace();
		}
		choosenColumns.forEach(k -> System.out.print(k + " "));
		System.out.println();
		if (!startDateColumn.equals("NA"))
			startDateIndex = Integer.parseInt(startDateColumn);
		if (!endDateColumn.equals("NA"))
			endDateIndex = Integer.parseInt(endDateColumn);
	}

	@SuppressWarnings("unchecked")
	@RequestMapping(method = RequestMethod.POST, path = "/postTrendColumns")
	public void postTrendColumns(String date, String dim) {

		dateIndex = Integer.parseInt(date);
		dimIndex = Integer.parseInt(dim);
	}

	@RequestMapping(method = RequestMethod.GET, path = "/doGeneralAnalysis")
	public GeneralAnalyticsModel doGeneralAnalysis() {
		long startTime = System.currentTimeMillis();
		System.out.println("General Analysis Started");

		List<SingleAnalyticsModel> singleColumnAnalytics = new ArrayList<SingleAnalyticsModel>();
		List<DualAnalyticsModel> dualColumnAnalytics = new ArrayList<DualAnalyticsModel>();
		List<ResolutionAnalyticsModel> resolutionAnalytics = new ArrayList<>();
		List<List<String>> combos = getAllCombinations();
		System.out.println("Combinations are fetched");
		combos.forEach(k -> {
			if (k.size() == 1) {
				singleColumnAnalytics.add(doAnalytics(k));
				System.out.println("Completed Single Analytics on " + k);
				if (startDateIndex > -1 && endDateIndex > -1) {
					resolutionAnalytics.add(doResolutionAnalytics(k));
					System.out.println("Completed Resolution analytics on " + k);
				}
			} else if (k.size() == 2) {
				dualColumnAnalytics.add(doDualAnalytics(k));
				System.out.println("Completed Dual Analytics on " + k);
			} else {
				System.out.println("More than 2 drill downs - So Ignoring");
			}
		});
		generalAnalyticsModel.setListOfSingleAnalyticsModel(singleColumnAnalytics);
		generalAnalyticsModel.setListOfDualAnalyticsModel(dualColumnAnalytics);
		generalAnalyticsModel.setListOfResolutionAnalyticsModel(resolutionAnalytics);
		System.out.println("General Analysis Completed in: " + (System.currentTimeMillis() - startTime) / 1000);
		return generalAnalyticsModel;
	}

	private ResolutionAnalyticsModel doResolutionAnalytics(List<String> choosenName) {

		ResolutionAnalyticsModel resolutionAnalyticsModel = new ResolutionAnalyticsModel();
		int colIndex = columnNames.get(choosenName.get(0));

		Map<String, List<Double>> resolutionAnalytics = new HashMap<>();
		Iterator<Row> iterator = datatypeSheet.iterator();
		iterator.next();
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();
			System.out.println("last fetched row: "+currentRow.getRowNum());
			Cell cell = currentRow.getCell(colIndex);
			Cell startDate = currentRow.getCell(startDateIndex);
			Cell endDate = currentRow.getCell(endDateIndex);

			String datatype = getDatatypeOfColumn(choosenName.get(0));
			if (null != endDate && endDate.toString().trim().length() > 0 && null != endDate.getDateCellValue()
					&& null != startDate && startDate.toString().trim().length() > 0
					&& null != startDate.getDateCellValue()) {
				long milliSecondsElapsed = endDate.getDateCellValue().getTime()
						- startDate.getDateCellValue().getTime();
				double diff = (double) milliSecondsElapsed / (1000 * 60 * 60 * 24);
//				double diff = TimeUnit.DAYS.convert(milliSecondsElapsed, TimeUnit.MILLISECONDS);
				switch (datatype) {
				case "String":
					if (cell != null && !cell.equals("")) {
						if (resolutionAnalytics.containsKey(cell.getStringCellValue())) {
							List<Double> existingList = resolutionAnalytics.get(cell.getStringCellValue());
							existingList.add(diff);
							resolutionAnalytics.put(cell.getStringCellValue(), existingList);
						} else {
							List<Double> newList = new ArrayList<>();
							newList.add(diff);
							resolutionAnalytics.put(cell.getStringCellValue(), newList);
						}
					}
					break;
				case "Boolean":
					if (cell != null && !cell.equals("")) {
						if (resolutionAnalytics.containsKey(String.valueOf(cell.getBooleanCellValue()))) {
							List<Double> existingList = resolutionAnalytics
									.get(String.valueOf(cell.getBooleanCellValue()));
							existingList.add(diff);
							resolutionAnalytics.put(String.valueOf(cell.getBooleanCellValue()), existingList);
						} else {
							List<Double> newList = new ArrayList<>();
							newList.add(diff);
							resolutionAnalytics.put(String.valueOf(cell.getBooleanCellValue()), newList);
						}
					}
					break;
				case "Numeric":
					if (cell != null && !cell.equals("")) {
						if (resolutionAnalytics.containsKey(String.valueOf(cell.getNumericCellValue()))) {
							List<Double> existingList = resolutionAnalytics
									.get(String.valueOf(cell.getNumericCellValue()));
							existingList.add(diff);
							resolutionAnalytics.put(String.valueOf(cell.getNumericCellValue()), existingList);
						} else {
							List<Double> newList = new ArrayList<>();
							newList.add(diff);
							resolutionAnalytics.put(String.valueOf(cell.getNumericCellValue()), newList);
						}
					}
					break;
				case "Date":
					if (cell != null && !cell.equals("")) {
						if (resolutionAnalytics.containsKey(String.valueOf(cell.getDateCellValue()))) {
							List<Double> existingList = resolutionAnalytics.get(cell.getStringCellValue());
							existingList.add(diff);
							resolutionAnalytics.put(String.valueOf(cell.getDateCellValue()), existingList);
						} else {
							List<Double> newList = new ArrayList<>();
							newList.add(diff);
							resolutionAnalytics.put(String.valueOf(cell.getDateCellValue()), newList);
						}
					}
					break;
				default:
					break;
				}
			}

		}
		Map<String, String> avgResolutionAnalytics = new HashMap<>();
		for (Map.Entry<String, List<Double>> entry : resolutionAnalytics.entrySet()) {
			double s[] = { 0 };
			entry.getValue().forEach(p -> {
				s[0] += p;
			});
			avgResolutionAnalytics.put(entry.getKey(), df2.format((double) s[0] / (entry.getValue().size())));
		}
		;
		avgResolutionAnalytics = sortMapByDValue(avgResolutionAnalytics);
		TreeSet<String> highlightedValues = new TreeSet<>();
		int i = 0;
		double resTotal = 0L;
		for (String value : avgResolutionAnalytics.values()) {
			resTotal += Double.parseDouble(value);
			if (i == 3 || Double.parseDouble(value) < 5) {
				break;
			}
			highlightedValues.add(value);
			i++;
		}

		resolutionAnalyticsModel.setModelName(choosenName.get(0));
		if (!(avgResolutionAnalytics.size() == 0))
			resolutionAnalyticsModel.setTotal(Double.parseDouble(df2.format(resTotal / avgResolutionAnalytics.size())));
		resolutionAnalyticsModel.setAnalytics(avgResolutionAnalytics);
		resolutionAnalyticsModel.setHighlightedValues(highlightedValues);
		return resolutionAnalyticsModel;
	}

	public List<List<String>> getAllCombinations() {
		List<List<String>> combinationList = new ArrayList<List<String>>();

		// Single combos
		choosenColumns.forEach(k -> {
			List<String> list = new ArrayList<String>();
			list.add(k);
			combinationList.add(list);
		});

		// Double combos
		for (int i = 0; i < choosenColumns.size(); i++) {
			for (int j = i + 1; j < choosenColumns.size(); j++) {
				List<String> list = new ArrayList<String>();
				list.add(choosenColumns.get(i));
				list.add(choosenColumns.get(j));
				combinationList.add(list);
			}
		}

		System.out.println(combinationList.size());

//	    for ( long i = 1; i < Math.pow(2, choosenColumns.size()); i++ ) {
//	        List<String> list = new ArrayList<String>();
//	        for ( int j = 0; j < choosenColumns.size(); j++ ) {
//	            if ( (i & (long) Math.pow(2, j)) > 0 ) {
//	                list.add(choosenColumns.get(j));
//	            }
//	        }
//	        combinationList.add(list);
//	    }
		return combinationList;
	}

	private SingleAnalyticsModel doAnalytics(List<String> choosenName) {

		SingleAnalyticsModel singleAnalyticsModel = new SingleAnalyticsModel();
		int colIndex = columnNames.get(choosenName.get(0));

		Map<String, Integer> columnAnalytics = new HashMap<>();
		Iterator<Row> iterator = datatypeSheet.iterator();
		iterator.next();
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();
			Cell cell = currentRow.getCell(colIndex);
			String datatype = getDatatypeOfColumn(choosenName.get(0));
			switch (datatype) {
			case "String":
				if (cell != null && !cell.equals("")) {
					if (columnAnalytics.containsKey(cell.getStringCellValue())) {
						columnAnalytics.put(cell.getStringCellValue(),
								columnAnalytics.get(cell.getStringCellValue()) + 1);
					} else {
						columnAnalytics.put(cell.getStringCellValue(), 1);
					}
				}
				break;
			case "Numeric":
				if (cell != null && !cell.equals("")) {
					if (columnAnalytics.containsKey(String.valueOf(cell.getNumericCellValue()))) {
						columnAnalytics.put(String.valueOf(cell.getNumericCellValue()),
								columnAnalytics.get(String.valueOf(cell.getNumericCellValue())) + 1);
					} else {
						columnAnalytics.put(String.valueOf(cell.getNumericCellValue()), 1);
					}
				}
				break;
			case "Boolean":
				if (cell != null && !cell.equals("")) {
					if (columnAnalytics.containsKey(String.valueOf(cell.getBooleanCellValue()))) {
						columnAnalytics.put(String.valueOf(cell.getBooleanCellValue()),
								columnAnalytics.get(String.valueOf(cell.getBooleanCellValue())) + 1);
					} else {
						columnAnalytics.put(String.valueOf(cell.getBooleanCellValue()), 1);
					}
				}
				break;
			case "Date":
				if (cell != null && !cell.equals("")) {
					if (columnAnalytics.containsKey(String.valueOf(cell.getDateCellValue()))) {
						columnAnalytics.put(String.valueOf(cell.getDateCellValue()),
								columnAnalytics.get(String.valueOf(cell.getDateCellValue())) + 1);
					} else {
						columnAnalytics.put(String.valueOf(cell.getDateCellValue()), 1);
					}
				}
				break;
			default:
				break;
			}

		}
		columnAnalytics = sortMapByValue(columnAnalytics);
		TreeSet<Integer> highlightedValues = new TreeSet<>();
		int i = 0;
		for (int value : columnAnalytics.values()) {
			if (i == 3 || ((double) value / (datatypeSheet.getLastRowNum()) * 100) < 5) {
				break;
			}
			highlightedValues.add(value);
			i++;
		}
		singleAnalyticsModel.setModelName(choosenName.get(0));
		singleAnalyticsModel.setTotal(datatypeSheet.getLastRowNum());
		singleAnalyticsModel.setAnalytics(columnAnalytics);
		singleAnalyticsModel.setHighlightedValues(highlightedValues);
		return singleAnalyticsModel;
	}

	private String getDatatypeOfColumn(String choosenName) {
		return columnTypes.get(choosenName);
	}

	private Map<String, Integer> sortMapByValue(Map<String, Integer> columnAnalytics) {
		final Map<String, Integer> sortedByCount = columnAnalytics.entrySet().stream()
				.sorted(Map.Entry.<String, Integer>comparingByValue().reversed())
				.collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue, (e1, e2) -> e1, LinkedHashMap::new));

		return sortedByCount;

	}

	private Map<String, String> sortMapByDValue(Map<String, String> columnAnalytics) {
		final Map<String, String> sortedByCount = columnAnalytics.entrySet().stream().sorted(
				(e1, e2) -> Double.compare(Double.parseDouble(e2.getValue()), Double.parseDouble(e1.getValue())))
				.collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue, (e1, e2) -> e1, LinkedHashMap::new));

		return sortedByCount;

	}

	@RequestMapping(method = RequestMethod.POST, path = "/doTrendAnalysis")
	public TrendAnalyticsModel performTrendAnalytics(String filter, String[] vals) {

		SortedMap<Integer, SortedMap<Integer, Integer>> trendDateAnalysis = new TreeMap<>();
		SortedMap<Integer, SortedMap<String, Integer>> trendDimAnalysis = new TreeMap<>();
		TreeSet<Integer> years = new TreeSet<Integer>();
		TreeSet<String> dims = new TreeSet<String>();

		List<TrendDateModel> trendDateModelList = new ArrayList<TrendDateModel>();
		List<TrendDimModel> trendDimModelList = new ArrayList<TrendDimModel>();

		// Filter details
		int filterIndex = -1;
		List<String> filterValues = new ArrayList<>();
		if (!filter.equals("No Filter")) {
			filterIndex = columnNames.get(filter);
			filterValues = Arrays.asList(vals);
		}

		Iterator<Row> iterator = datatypeSheet.iterator();
		iterator.next();
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();
			Cell dateCell = currentRow.getCell(dateIndex);
			Cell dimCell = currentRow.getCell(dimIndex);
			Cell filterCell = null;
			if (filterIndex > 0)
				filterCell = currentRow.getCell(filterIndex);
			if (null == filterCell || (filterValues.contains((filterCell.getStringCellValue())))) {
				if (dateCell != null &&  dateCell.toString().trim().length() > 0 && null != dateCell.getDateCellValue()) {

					LocalDate a = dateCell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
					// System.out.println(a.getYear() + " " + a.getMonth().getValue());
					int monthName = a.getMonth().getValue();
					if (!years.contains(a.getYear())) {
						years.add(a.getYear());
					}
					if (trendDateAnalysis.containsKey(monthName)) {//
						SortedMap<Integer, Integer> innerMap = trendDateAnalysis.get(monthName);//
						if (innerMap.containsKey(a.getYear())) {
							innerMap.put(a.getYear(), innerMap.get(a.getYear()) + 1);
						} else {
							innerMap.put(a.getYear(), 1);
						}
						trendDateAnalysis.put(monthName, innerMap);//
					} else {
						SortedMap<Integer, Integer> innerMap = new TreeMap<Integer, Integer>();
						innerMap.put(a.getYear(), 1);
						trendDateAnalysis.put(monthName, innerMap);//
					}
					String dimVal = dimCell.getStringCellValue();
					if (!dims.contains(dimVal)) {
						dims.add(dimVal);
					}
					if (trendDimAnalysis.containsKey(a.getYear())) {
						SortedMap<String, Integer> innerMap = trendDimAnalysis.get(a.getYear());
						if (innerMap.containsKey(dimVal)) {
							innerMap.put(dimVal, innerMap.get(dimVal) + 1);
						} else {
							innerMap.put(dimVal, 1);
						}
						trendDimAnalysis.put(a.getYear(), innerMap);//
					} else {
						SortedMap<String, Integer> innerMap = new TreeMap<String, Integer>();
						innerMap.put(dimVal, 1);
						trendDimAnalysis.put(a.getYear(), innerMap);//
					}
				}
			}
		}
		trendDateAnalysis.forEach((k, v) -> {
			years.forEach(z -> {
				if (!v.containsKey(z)) {
					v.put(z, 0);
				}
			});
			TrendDateModel dateModel = new TrendDateModel();
			List<Integer> list = new ArrayList<Integer>();
			dateModel.setMonth(new DateFormatSymbols().getShortMonths()[k - 1]);
			v.forEach((b, c) -> list.add(c));
			dateModel.setCountList(list);

			trendDateModelList.add(dateModel);
		});

		trendDimAnalysis.forEach((l, m) -> {
			dims.forEach(n -> {
				if (!m.containsKey(n)) {
					m.put(n, 0);
				}
			});
			TrendDimModel dimModel = new TrendDimModel();
			List<Integer> list = new ArrayList<Integer>();
			dimModel.setDim(l);
			m.forEach((b, c) -> list.add(c));
			dimModel.setCount(list);

			trendDimModelList.add(dimModel);
		});

		TrendAnalyticsModel trendModel = new TrendAnalyticsModel();
		trendModel.setTrendDateModelList(trendDateModelList);
		trendModel.setTrendDimModelList(trendDimModelList);
		TreeSet<String> yearSet = new TreeSet<String>();
		years.forEach(y -> yearSet.add(String.valueOf(y)));
		trendModel.setYears(yearSet);
		TreeSet<String> dimSet = new TreeSet<String>();
		dims.forEach(p -> dimSet.add(String.valueOf(p)));
		trendModel.setDims(dimSet);
		trendModel.setAllDims(allDims);

		return trendModel;
	}

	@RequestMapping(method = RequestMethod.POST, path = "/exportTextData")
	public ResponseEntity<?> exportTextData(@RequestBody TextAnalyticsModel[] listOfLists) {

		ByteArrayOutputStream fileOut = new ByteArrayOutputStream();
		try {
			System.out.println("Started Text Export");
			// Create a Workbook
			Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

			/*
			 * CreationHelper helps us create instances of various things like DataFormat,
			 * Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way
			 */
			CreationHelper createHelper = workbook.getCreationHelper();

			// Create a Sheet
			Sheet sheet = workbook.createSheet("Text Analytics");
			sheet.setDisplayGridlines(false);
			sheet.setZoom(80);

			// Create a Font for styling header cells
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short) 11);
			headerFont.setColor(IndexedColors.WHITE.getIndex());

			// Create a Font for styling header cells
			Font rowFont = workbook.createFont();
			rowFont.setFontHeightInPoints((short) 11);
			rowFont.setColor(IndexedColors.BLACK.getIndex());

			// Create a Font for styling header cells
			Font rowFontWhite = workbook.createFont();
			rowFontWhite.setBold(true);
			rowFontWhite.setFontHeightInPoints((short) 11);
			rowFontWhite.setColor(IndexedColors.BLACK.getIndex());

			// Create a CellStyle with the font
			CellStyle headerCellStyle = workbook.createCellStyle();
			headerCellStyle.setFont(headerFont);
			headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
			headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerCellStyle.setFillForegroundColor(IndexedColors.BLACK.getIndex());
			headerCellStyle.setBorderBottom(BorderStyle.MEDIUM);
			headerCellStyle.setBorderTop(BorderStyle.MEDIUM);
			headerCellStyle.setBorderRight(BorderStyle.MEDIUM);
			headerCellStyle.setBorderLeft(BorderStyle.MEDIUM);

			// Create a CellStyle with the font
			CellStyle headerCellStyleBlue = workbook.createCellStyle();
			headerCellStyleBlue.setFont(headerFont);
			headerCellStyleBlue.setAlignment(HorizontalAlignment.CENTER);
			headerCellStyleBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerCellStyleBlue.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());

			// Create a CellStyle with the font
			CellStyle headerCellStyleGreen = workbook.createCellStyle();
			headerCellStyleGreen.setFont(headerFont);
			headerCellStyleGreen.setAlignment(HorizontalAlignment.CENTER);
			headerCellStyleGreen.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerCellStyleGreen.setFillForegroundColor(IndexedColors.GREEN.getIndex());

			CellStyle yellowRowCellStyle = workbook.createCellStyle();
			yellowRowCellStyle.setFont(rowFontWhite);
			yellowRowCellStyle.setAlignment(HorizontalAlignment.CENTER);
			yellowRowCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			yellowRowCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			yellowRowCellStyle.setBorderBottom(BorderStyle.THIN);
			yellowRowCellStyle.setBorderTop(BorderStyle.THIN);
			yellowRowCellStyle.setBorderRight(BorderStyle.THIN);
			yellowRowCellStyle.setBorderLeft(BorderStyle.THIN);

			// Create a CellStyle with the font
			CellStyle rowCellStyle = workbook.createCellStyle();
			rowCellStyle.setFont(rowFont);
			rowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			rowCellStyle.setBorderBottom(BorderStyle.THIN);
			rowCellStyle.setBorderTop(BorderStyle.THIN);
			rowCellStyle.setBorderRight(BorderStyle.THIN);
			rowCellStyle.setBorderLeft(BorderStyle.THIN);

			int rowIndex = 0;
			Row hrow = sheet.getRow(rowIndex);
			if (hrow == null) {
				hrow = sheet.createRow(rowIndex);
			}
			CellRangeAddress hmergedRegion = new CellRangeAddress(rowIndex, rowIndex, 0, 2); // row 1 col B and C
			sheet.addMergedRegion(hmergedRegion);
			Cell hcell = hrow.createCell(0);
			hcell.setCellValue("Overall Stats");
			hcell.setCellStyle(headerCellStyleGreen);
			rowIndex++;

			Row hrow1 = sheet.getRow(rowIndex);
			if (hrow1 == null) {
				hrow1 = sheet.createRow(rowIndex);
			}
			CellRangeAddress hmergedRegion1 = new CellRangeAddress(rowIndex, rowIndex, 0, 2); // row 1 col B and C
			sheet.addMergedRegion(hmergedRegion1);
			Cell hcell1 = hrow1.createCell(0);
			hcell1.setCellValue("Breakdown of data by top issues");
			hcell1.setCellStyle(headerCellStyleBlue);
			rowIndex++;

			// Loop through the list and populate the excel

			// Create cells
			Row headerRow = sheet.getRow(rowIndex);
			if (headerRow == null) {
				headerRow = sheet.createRow(rowIndex);
			}
			String[] columns = { "Request/Issue", "Count", "Percentage" };
			for (int i = 0; i < columns.length; i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(columns[i]);
				cell.setCellStyle(headerCellStyle);
			}

			rowIndex++;
			int total = 0;
			int totalPer = 0;
			for (TextAnalyticsModel list : listOfLists) {
				Row row = sheet.getRow(rowIndex);
				if (null == row)
					row = sheet.createRow(rowIndex);
				CellStyle cs = rowCellStyle;
				Cell cell1 = row.createCell(0);

				String cellValue = list.getListName() + ": Issues Related to ";
				for (String data : list.getData()) {
					cellValue += data.split("-")[0] + ",";
				}
				cellValue += " etc.,";
				cell1.setCellValue(cellValue);
				cell1.setCellStyle(rowCellStyle);
				Cell cell2 = row.createCell(1);
				cell2.setCellValue(list.getHeaderTotal());

				total += list.getHeaderTotal();

				BigDecimal bd = new BigDecimal((double) list.getHeaderTotal() / (datatypeSheet.getLastRowNum()) * 100)
						.setScale(2, RoundingMode.HALF_UP);
				cell2.setCellStyle(rowCellStyle);
				Cell cell3 = row.createCell(2);
				cell3.setCellValue(bd.doubleValue());
				cell3.setCellStyle(rowCellStyle);

				totalPer += list.getHeaderPer();
				rowIndex++;
			}

			// Total Row
			Row totalRow = sheet.getRow(rowIndex);
			if (totalRow == null) {
				totalRow = sheet.createRow(rowIndex);
			}
			Cell cellh = totalRow.createCell(0);
			cellh.setCellValue("Total");
			cellh.setCellStyle(yellowRowCellStyle);

			Cell cellh1 = totalRow.createCell(1);
			cellh1.setCellValue(total);
			cellh1.setCellStyle(yellowRowCellStyle);

			Cell cellh2 = totalRow.createCell(2);
			cellh2.setCellValue(totalPer);
			cellh2.setCellStyle(yellowRowCellStyle);
			rowIndex++;
			rowIndex++;
			rowIndex++;

			Row headerRow2 = sheet.getRow(rowIndex);
			if (headerRow2 == null) {
				headerRow2 = sheet.createRow(rowIndex);
			}
			CellRangeAddress hmergedRegion2 = new CellRangeAddress(rowIndex, rowIndex, 0, 2); // row 1 col B and C
			sheet.addMergedRegion(hmergedRegion2);
			Cell hcell2 = headerRow2.createCell(0);
			hcell2.setCellValue("Break down of second level analytics");
			hcell2.setCellStyle(headerCellStyleBlue);
			rowIndex++;

			Row headerRow3 = sheet.getRow(rowIndex);
			if (headerRow3 == null) {
				headerRow3 = sheet.createRow(rowIndex);
			}

			for (int i = 0; i < columns.length; i++) {
				Cell cell = headerRow3.createCell(i);
				cell.setCellValue(columns[i]);
				cell.setCellStyle(headerCellStyle);
			}

			rowIndex++;
			for (TextAnalyticsModel list : listOfLists) {
				Row row = sheet.getRow(rowIndex);
				if (null == row)
					row = sheet.createRow(rowIndex);
				Cell cell1 = row.createCell(0);
				cell1.setCellValue(list.getListName());
				cell1.setCellStyle(yellowRowCellStyle);

				Cell cell2 = row.createCell(1);
				cell2.setCellValue(list.getHeaderTotal());
				cell2.setCellStyle(yellowRowCellStyle);

				Cell cell3 = row.createCell(2);
				cell3.setCellValue(list.getHeaderPer());
				cell3.setCellStyle(yellowRowCellStyle);

				rowIndex++;

				for (String data : list.getData()) {
					String[] dataSplit = data.split("-");

					Row row1 = sheet.getRow(rowIndex);
					if (null == row1)
						row1 = sheet.createRow(rowIndex);
					Cell cell11 = row1.createCell(0);
					cell11.setCellValue("Issues related to: " + dataSplit[0]);
					cell11.setCellStyle(rowCellStyle);

					Cell cell12 = row1.createCell(1);
					cell12.setCellValue(dataSplit[1]);
					cell12.setCellStyle(rowCellStyle);

					BigDecimal bd = new BigDecimal(
							(double) Integer.parseInt(dataSplit[1]) / (list.getHeaderTotal()) * 100).setScale(2,
									RoundingMode.HALF_UP);
					Cell cell13 = row1.createCell(2);
					cell13.setCellValue(bd.doubleValue());
					cell13.setCellStyle(rowCellStyle);

					rowIndex++;
				}
			}

			// Resize all columns to fit the content size
			for (int i = 0; i < 3; i++) {
				sheet.autoSizeColumn(i);
			}

			// Write the output to a file
			workbook.write(fileOut);
			fileOut.close();

			// Closing the workbook
			workbook.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(MediaType.parseMediaType("application/vnd.ms-excel"));
		System.out.println("Completed Text Export");
		return new ResponseEntity<byte[]>(fileOut.toByteArray(), headers, HttpStatus.OK);
	}

	@RequestMapping(method = RequestMethod.GET, path = "/exportData")
	public ResponseEntity<?> exportData(@RequestParam String type) {
		String[] columns = { "Sl.No.", "", "Count", "Percentage" };
		String[] columns1 = { "Sl.No.", "", "", "Count", "Percentage" };
		String[] columns2 = { "Sl.No.", "", "Average Resolution (In Days)" };
		GeneralAnalyticsModel generalAnalysisModel = generalAnalyticsModel;
		ByteArrayOutputStream fileOut = new ByteArrayOutputStream();
		try {
			System.out.println("Started Export");
			// Create a Workbook
			Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

			/*
			 * CreationHelper helps us create instances of various things like DataFormat,
			 * Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way
			 */
			CreationHelper createHelper = workbook.getCreationHelper();

			// Create a Sheet
			Sheet sheet = workbook.createSheet("General Analytics_L1");
			sheet.setDisplayGridlines(false);
			sheet.setZoom(70);

			// Create a Font for styling header cells
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short) 10);
			headerFont.setColor(IndexedColors.WHITE.getIndex());

			// Create a CellStyle with the font
			CellStyle headerCellStyle = workbook.createCellStyle();
			headerCellStyle.setFont(headerFont);
			headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerCellStyle.setFillForegroundColor(IndexedColors.BLACK.getIndex());
			headerCellStyle.setBorderBottom(BorderStyle.MEDIUM);
			headerCellStyle.setBorderTop(BorderStyle.MEDIUM);
			headerCellStyle.setBorderRight(BorderStyle.MEDIUM);
			headerCellStyle.setBorderLeft(BorderStyle.MEDIUM);

			// Create a Font for styling header cells
			Font rowFont = workbook.createFont();
			rowFont.setFontHeightInPoints((short) 8);
			rowFont.setColor(IndexedColors.BLACK.getIndex());

			// Create a CellStyle with the font
			CellStyle rowCellStyle = workbook.createCellStyle();
			rowCellStyle.setFont(rowFont);
			rowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			rowCellStyle.setBorderBottom(BorderStyle.THIN);
			rowCellStyle.setBorderTop(BorderStyle.THIN);
			rowCellStyle.setBorderRight(BorderStyle.THIN);
			rowCellStyle.setBorderLeft(BorderStyle.THIN);

			// Create a CellStyle with the font
			CellStyle yellowRowCellStyle = workbook.createCellStyle();
			yellowRowCellStyle.setFont(rowFont);
			yellowRowCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			yellowRowCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			yellowRowCellStyle.setBorderBottom(BorderStyle.THIN);
			yellowRowCellStyle.setBorderTop(BorderStyle.THIN);
			yellowRowCellStyle.setBorderRight(BorderStyle.THIN);
			yellowRowCellStyle.setBorderLeft(BorderStyle.THIN);

			// Create a CellStyle with the font
			CellStyle totalCellStyle = workbook.createCellStyle();
			totalCellStyle.setFont(headerFont);
			totalCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			totalCellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
			totalCellStyle.setBorderBottom(BorderStyle.THIN);
			totalCellStyle.setBorderTop(BorderStyle.THIN);
			totalCellStyle.setBorderRight(BorderStyle.THIN);
			totalCellStyle.setBorderLeft(BorderStyle.THIN);
			totalCellStyle.setAlignment(HorizontalAlignment.CENTER);

			Font thFont = workbook.createFont();
			headerFont.setBold(true);
			rowFont.setFontHeightInPoints((short) 10);
			rowFont.setColor(IndexedColors.BLACK.getIndex());
			// Create a CellStyle with the font
			CellStyle thstyle = workbook.createCellStyle();
			thstyle.setFont(thFont);
			thstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			thstyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
			thstyle.setBorderBottom(BorderStyle.THIN);
			thstyle.setBorderTop(BorderStyle.THIN);
			thstyle.setBorderRight(BorderStyle.THIN);
			thstyle.setBorderLeft(BorderStyle.THIN);
			thstyle.setAlignment(HorizontalAlignment.CENTER);

			// Create a Row

			List<SingleAnalyticsModel> sAnalysisModel = generalAnalysisModel.getListOfSingleAnalyticsModel();
			int columnNum[] = { 0 };
			int rowNum[] = { 2 };

			for (SingleAnalyticsModel singleAnalyticsModel : sAnalysisModel) {
				columns[1] = singleAnalyticsModel.getModelName();
				Map<String, Integer> analysisModel = singleAnalyticsModel.getAnalytics();
				TreeSet<Integer> toBeHighlighted = singleAnalyticsModel.getHighlightedValues();
				Row hrow = sheet.getRow(0);
				if (hrow == null) {
					hrow = sheet.createRow(0);
				}
				CellRangeAddress hmergedRegion = new CellRangeAddress(0, 0, columnNum[0], columnNum[0] + 3); // row 1
																												// col B
																												// and C
				sheet.addMergedRegion(hmergedRegion);
				Cell hcell = hrow.createCell(columnNum[0]);
				hcell.setCellValue(type + " By " + singleAnalyticsModel.getModelName());
				hcell.setCellStyle(thstyle);

				// Create cells
				Row headerRow = sheet.getRow(1);
				if (headerRow == null) {
					headerRow = sheet.createRow(1);
				}
				for (int i = 0; i < columns.length; i++) {
					Cell cell = headerRow.createCell(columnNum[0] + i);
					cell.setCellValue(columns[i]);
					cell.setCellStyle(headerCellStyle);
				}

				analysisModel.forEach((k, v) -> {
					Row row = sheet.getRow(rowNum[0]);
					if (null == row) {
						row = sheet.createRow(rowNum[0]);
					}
					CellStyle cs = toBeHighlighted.contains(v) ? yellowRowCellStyle : rowCellStyle;
					Cell cell1 = row.createCell(columnNum[0]);
					cell1.setCellValue(-1 + rowNum[0]++);
					cell1.setCellStyle(cs);
					Cell cell2 = row.createCell(columnNum[0] + 1);
					cell2.setCellValue(k);
					cell2.setCellStyle(cs);
					Cell cell3 = row.createCell(columnNum[0] + 2);
					cell3.setCellValue(v);
					cell3.setCellStyle(cs);

					BigDecimal bd = new BigDecimal((double) v / (singleAnalyticsModel.getTotal()) * 100).setScale(2,
							RoundingMode.HALF_UP);
					Cell cell4 = row.createCell(columnNum[0] + 3);
					cell4.setCellValue(bd.doubleValue());
					cell4.setCellStyle(cs);
				});
				Row row = sheet.getRow(rowNum[0]);
				if (null == row) {
					row = sheet.createRow(rowNum[0]);
				}
				CellRangeAddress mergedRegion = new CellRangeAddress(rowNum[0], rowNum[0], columnNum[0],
						columnNum[0] + 1); // row 1 col B and C
				sheet.addMergedRegion(mergedRegion);
				Cell cell1 = row.createCell(columnNum[0]);
				cell1.setCellValue("Grand Total");
				cell1.setCellStyle(totalCellStyle);
				Cell cell3 = row.createCell(columnNum[0] + 2);
				cell3.setCellValue(singleAnalyticsModel.getTotal());
				cell3.setCellStyle(totalCellStyle);
				BigDecimal bd = new BigDecimal(
						(double) singleAnalyticsModel.getTotal() / (singleAnalyticsModel.getTotal()) * 100).setScale(2,
								RoundingMode.HALF_UP);
				Cell cell4 = row.createCell(columnNum[0] + 3);
				cell4.setCellValue(bd.doubleValue());
				cell4.setCellStyle(totalCellStyle);
				columnNum[0] = columnNum[0] + 5;
				rowNum[0] = 2;
			}

			// Resize all columns to fit the content size
			for (int i = 0; i < columnNum[0]; i++) {
				sheet.autoSizeColumn(i);
			}
			System.out.println("Completed L1");

			// Double drill down
			// Create a Sheet
			Sheet sheet1 = workbook.createSheet("General Analytics_L2");
			sheet1.setDisplayGridlines(false);
			sheet1.setZoom(70);

			// Create a Row

			List<DualAnalyticsModel> dAnalysisModel = generalAnalysisModel.getListOfDualAnalyticsModel();
			int columnNum1[] = { 0 };
			int rowNum1[] = { 2 };

			for (DualAnalyticsModel dualAnalyticsModel : dAnalysisModel) {
				String[] model = dualAnalyticsModel.getModelName().split("-");
				columns1[1] = model[0];
				columns1[2] = model[1];

				Map<String, Map<String, Integer>> analysisModel = dualAnalyticsModel.getDualAnalytics();
				TreeSet<Integer> toBeHighlighted = dualAnalyticsModel.getHighlightedValues();
				Row hrow = sheet1.getRow(0);
				if (null == hrow) {
					hrow = sheet1.createRow(0);
				}
				CellRangeAddress hmergedRegion = new CellRangeAddress(0, 0, columnNum1[0], columnNum1[0] + 4); // row 1
																												// col B
																												// and C
				sheet1.addMergedRegion(hmergedRegion);
				Cell hcell = hrow.createCell(columnNum1[0]);
				hcell.setCellValue(type + " By " + model[0] + " and " + model[1]);
				hcell.setCellStyle(thstyle);

				Row headerRow1 = sheet1.getRow(1);
				if (headerRow1 == null) {
					headerRow1 = sheet1.createRow(1);
				}
				// Create cells
				for (int i = 0; i < columns1.length; i++) {
					Cell cell = headerRow1.createCell(columnNum1[0] + i);
					cell.setCellValue(columns1[i]);
					cell.setCellStyle(headerCellStyle);
				}
				int rowStart[] = { rowNum1[0] };
				analysisModel.forEach((k, v) -> {
					v.forEach((a, b) -> {
						Row row = sheet1.getRow(rowNum1[0]);
						if (null == row)
							row = sheet1.createRow(rowNum1[0]);
						CellStyle cs = ((toBeHighlighted.contains(b))
								&& (((double) b / dualAnalyticsModel.getTotal()) * 100) > 5) ? yellowRowCellStyle
										: rowCellStyle;
						Cell cell1 = row.createCell(columnNum1[0]);
						cell1.setCellValue(-1 + rowNum1[0]++);
						cell1.setCellStyle(rowCellStyle);
						Cell cell2 = row.createCell(columnNum1[0] + 1);
						cell2.setCellValue(k);
						cell2.setCellStyle(rowCellStyle);
						Cell cell3 = row.createCell(columnNum1[0] + 2);
						cell3.setCellValue(a);
						cell3.setCellStyle(cs);
						Cell cell4 = row.createCell(columnNum1[0] + 3);
						cell4.setCellValue(b);
						cell4.setCellStyle(cs);

						BigDecimal bd = new BigDecimal((double) b / (dualAnalyticsModel.getTotal()) * 100).setScale(2,
								RoundingMode.HALF_UP);
						Cell cell5 = row.createCell(columnNum1[0] + 4);
						cell5.setCellValue(bd.doubleValue());
						cell5.setCellStyle(cs);
					});
					int endRow = rowStart[0] + v.size() - 1;
					if (v.size() > 1) {

						CellRangeAddress mergedRegion = new CellRangeAddress(rowStart[0], endRow, columnNum1[0] + 1,
								columnNum1[0] + 1); // row 1 col B and C
						sheet1.addMergedRegion(mergedRegion);
					}
					rowStart[0] = endRow + 1;
				});
				Row row = sheet1.getRow(rowNum1[0]);
				if (null == row)
					row = sheet1.createRow(rowNum1[0]);

				Cell cell1 = row.createCell(columnNum1[0]);
				cell1.setCellValue("Grand Total");
				cell1.setCellStyle(totalCellStyle);
				CellRangeAddress mergedRegion = new CellRangeAddress(rowNum1[0], rowNum1[0], columnNum1[0],
						columnNum1[0] + 2); // row 1 col B and C
				sheet1.addMergedRegion(mergedRegion);
				Cell cell4 = row.createCell(columnNum1[0] + 3);
				cell4.setCellValue(dualAnalyticsModel.getTotal());
				cell4.setCellStyle(totalCellStyle);
				BigDecimal bd = new BigDecimal(
						(double) dualAnalyticsModel.getTotal() / (datatypeSheet.getLastRowNum()) * 100).setScale(2,
								RoundingMode.HALF_UP);
				Cell cell5 = row.createCell(columnNum1[0] + 4);
				cell5.setCellValue(bd.doubleValue());
				cell5.setCellStyle(totalCellStyle);
				columnNum1[0] = columnNum1[0] + 6;
				rowNum1[0] = 2;
			}

			// Resize all columns to fit the content size
			for (int i = 0; i < columnNum1[0]; i++) {
				sheet1.autoSizeColumn(i);
			}
			System.out.println("Completed L2");
			// Resolution Analytics
			Sheet sheet2 = workbook.createSheet("Resolution Analytics");
			sheet2.setDisplayGridlines(false);
			sheet2.setZoom(70);

			List<ResolutionAnalyticsModel> resAnalyticsModel = generalAnalysisModel.getListOfResolutionAnalyticsModel();
			int columnNum2[] = { 0 };
			int rowNum2[] = { 2 };

			for (ResolutionAnalyticsModel resAnalytics : resAnalyticsModel) {
				columns2[1] = resAnalytics.getModelName();
				Map<String, String> analysisModel = resAnalytics.getAnalytics();
				TreeSet<String> toBeHighlighted = resAnalytics.getHighlightedValues();
				Row hrow = sheet2.getRow(0);
				if (hrow == null) {
					hrow = sheet2.createRow(0);
				}
				CellRangeAddress hmergedRegion = new CellRangeAddress(0, 0, columnNum2[0], columnNum2[0] + 2); // row 1
																												// col B
																												// and C
				sheet2.addMergedRegion(hmergedRegion);
				Cell hcell = hrow.createCell(columnNum2[0]);
				hcell.setCellValue(type + " By " + resAnalytics.getModelName());
				hcell.setCellStyle(thstyle);

				// Create cells
				Row headerRow = sheet2.getRow(1);
				if (headerRow == null) {
					headerRow = sheet2.createRow(1);
				}
				for (int i = 0; i < columns2.length; i++) {
					Cell cell = headerRow.createCell(columnNum2[0] + i);
					cell.setCellValue(columns2[i]);
					cell.setCellStyle(headerCellStyle);
				}

				analysisModel.forEach((k, v) -> {
					Row row = sheet2.getRow(rowNum2[0]);
					if (null == row) {
						row = sheet2.createRow(rowNum2[0]);
					}
					CellStyle cs = toBeHighlighted.contains(v) ? yellowRowCellStyle : rowCellStyle;
					Cell cell1 = row.createCell(columnNum2[0]);
					cell1.setCellValue(-1 + rowNum2[0]++);
					cell1.setCellStyle(cs);
					Cell cell2 = row.createCell(columnNum2[0] + 1);
					cell2.setCellValue(k);
					cell2.setCellStyle(cs);
					BigDecimal bd = new BigDecimal(v).setScale(2, RoundingMode.HALF_UP);
					Cell cell3 = row.createCell(columnNum2[0] + 2);
					cell3.setCellValue(bd.doubleValue());
					cell3.setCellStyle(cs);
				});
				Row row = sheet2.getRow(rowNum2[0]);
				if (null == row) {
					row = sheet2.createRow(rowNum2[0]);
				}
				CellRangeAddress mergedRegion = new CellRangeAddress(rowNum2[0], rowNum2[0], columnNum2[0],
						columnNum2[0] + 1); // row 1 col B and C
				sheet2.addMergedRegion(mergedRegion);
				Cell cell1 = row.createCell(columnNum2[0]);
				cell1.setCellValue("Grand Total");
				cell1.setCellStyle(totalCellStyle);
				Cell cell3 = row.createCell(columnNum2[0] + 2);
				cell3.setCellValue(resAnalytics.getTotal());
				cell3.setCellStyle(totalCellStyle);
				columnNum2[0] = columnNum2[0] + 4;
				rowNum2[0] = 2;
			}

			// Resize all columns to fit the content size
			for (int i = 0; i < columnNum2[0]; i++) {
				sheet2.autoSizeColumn(i);
			}

			// Write the output to a file

			workbook.write(fileOut);
			fileOut.close();

			// Closing the workbook
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(MediaType.parseMediaType("application/vnd.ms-excel"));
		System.out.println("Completed Everything");
		return new ResponseEntity<byte[]>(fileOut.toByteArray(), headers, HttpStatus.OK);
	}

	@RequestMapping(method = RequestMethod.POST, path = "/fetchAllCombinations")
	public void fetchAllCombinations(String stopWords, String noOfCombos, String[] combos, String fileName) {
		System.out.println("File Name: " + fileName);
		System.out.println("No. Of Combinations needed: " + noOfCombos);
		System.out.println("Combinations of words: " + combos);
		System.out.println("Stop Words: " + stopWords);

	}

}
