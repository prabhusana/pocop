/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

 */
package org.zkoss.zss.app.ui.menu;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.Date;
import java.util.Iterator;

import org.apache.commons.codec.binary.StringUtils;
import org.zkoss.poi.xssf.usermodel.XSSFCell;
import org.zkoss.poi.xssf.usermodel.XSSFRow;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.select.Selectors;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.app.repository.BookRepository;
import org.zkoss.zss.app.repository.BookRepositoryFactory;
import org.zkoss.zss.app.ui.AppEvts;
import org.zkoss.zss.app.ui.CtrlBase;
import org.zkoss.zss.app.ui.UiUtil;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.Version;
import org.zkoss.zss.ui.sys.UndoableActionManager;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menubar;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Messagebox;

import tutorial.MyProperties;

/**
 * 
 * @author dennis
 * 
 */
public class MainMenubarCtrl extends CtrlBase<Menubar> {

	public MainMenubarCtrl() {
		super(true);
	}

	@Wire
	Menuitem newFile;
	@Wire
	Menuitem openManageFile;
	@Wire
	Menuitem saveFile;
	@Wire
	Menuitem saveFileAs;
	@Wire
	Menuitem saveFileAndClose;
	@Wire
	Menuitem closeFile;
	@Wire
	Menuitem exportFile;
	@Wire
	Menuitem exportPdf;

	@Wire
	Menuitem undo;
	@Wire
	Menuitem redo;

	@Wire
	Menuitem toggleFormulaBar;
	@Wire
	Menuitem freezePanel;
	@Wire
	Menuitem unfreezePanel;
	@Wire
	Menu freezeRows;
	@Wire
	Menu freezeCols;

	@Wire
	Menuitem saveDb;
	@Wire
	Menuitem PreviousCOP;
	@Wire
	Menuitem VendorDetails;

	String bookname;
	String copValue;

	protected void onAppEvent(String event, Object data) {
		if (AppEvts.ON_CHANGED_SPREADSHEET.equals(event)) {
			doUpdateMenu((Spreadsheet) data);
		} else if (AppEvts.ON_UPDATE_UNDO_REDO.equals(event)) {
			doUpdateMenu((Spreadsheet) data);
		}

	}

	private void doUpdateMenu(Spreadsheet sparedsheet) {

		boolean hasBook = sparedsheet.getBook() != null;
		bookname = sparedsheet.getBook().getBookName();
		System.out.println("bookname from menubarctrl " + bookname);
		boolean isEE = "EE".equals(Version.getEdition());
		// new and open are always on
		newFile.setDisabled(false);
		openManageFile.setDisabled(false);
		boolean readonly = UiUtil.isRepositoryReadonly();
		boolean disabled = !hasBook;
		saveFile.setDisabled(disabled || readonly);
		saveFileAs.setDisabled(disabled || readonly);
		saveFileAndClose.setDisabled(disabled || readonly);
		closeFile.setDisabled(disabled);
		exportFile.setDisabled(disabled);
		exportPdf.setDisabled(!isEE || disabled);

		UndoableActionManager uam = sparedsheet.getUndoableActionManager();

		String label = Labels.getLabel("zssapp.mainMenu.edit.undo");
		if (isEE && uam.isUndoable()) {
			undo.setDisabled(false);
			label = label + ":" + uam.getUndoLabel();
		} else {
			undo.setDisabled(true);
		}
		undo.setLabel(label);

		label = Labels.getLabel("zssapp.mainMenu.edit.redo");
		if (isEE && uam.isRedoable()) {
			redo.setDisabled(false);
			label = label + ":" + uam.getRedoLabel();
		} else {
			redo.setDisabled(true);
		}
		redo.setLabel(label);

		// toggleFormulaBar.setDisabled(disabled); //don't need to care the book
		// load or not.
		toggleFormulaBar.setChecked(sparedsheet.isShowFormulabar());

		freezePanel.setDisabled(!isEE || disabled);
		Sheet sheet = sparedsheet.getSelectedSheet();
		unfreezePanel.setDisabled(!isEE || disabled || !(sheet.getRowFreeze() > 0 || sheet.getColumnFreeze() > 0));

		for (Component comp : Selectors.find(freezeRows, "menuitem")) {
			((Menuitem) comp).setDisabled(!isEE || disabled);
		}
		for (Component comp : Selectors.find(freezeCols, "menuitem")) {
			((Menuitem) comp).setDisabled(!isEE || disabled);
		}
	}

	@Listen("onClick=#newFile")
	public void onNew() {
		pushAppEvent(AppEvts.ON_NEW_BOOK);
	}

	@Listen("onClick=#openManageFile")
	public void onOpen() {
		pushAppEvent(AppEvts.ON_OPEN_MANAGE_BOOK);
	}

	@Listen("onClick=#saveFile")
	public void onSave() {
		pushAppEvent(AppEvts.ON_SAVE_BOOK);
	}

	@Listen("onClick=#saveFileAs")
	public void onSaveAs() {
		pushAppEvent(AppEvts.ON_SAVE_BOOK_AS);
	}

	@Listen("onClick=#saveFileAndClose")
	public void onSaveClose() {
		pushAppEvent(AppEvts.ON_SAVE_CLOSE_BOOK);
	}

	@Listen("onClick=#closeFile")
	public void onClose() {
		pushAppEvent(AppEvts.ON_CLOSE_BOOK);
	}

	@Listen("onClick=#exportFile")
	public void onExport() {
		pushAppEvent(AppEvts.ON_EXPORT_BOOK);
	}

	@Listen("onClick=#exportPdf")
	public void onExportPdf() {
		pushAppEvent(AppEvts.ON_EXPORT_BOOK_PDF);
	}

	@Listen("onToggleFormulaBar=#mainMenubar")
	public void onToggleFormulaBar() {
		pushAppEvent(AppEvts.ON_TOGGLE_FORMULA_BAR);
	}

	@Listen("onFreezePanel=#mainMenubar")
	public void onFreezePanel() {
		pushAppEvent(AppEvts.ON_FREEZE_PNAEL);
	}

	@Listen("onUnfreezePanel=#mainMenubar")
	public void onUnfreezePanel() {
		pushAppEvent(AppEvts.ON_UNFREEZE_PANEL);
	}

	@Listen("onViewFreezeRows=#mainMenubar")
	public void onViewFreezeRows(ForwardEvent event) {
		int index = Integer.parseInt((String) event.getData());
		pushAppEvent(AppEvts.ON_FREEZE_ROW, index);
	}

	@Listen("onViewFreezeCols=#mainMenubar")
	public void onViewFreezeCols(ForwardEvent event) {
		int index = Integer.parseInt((String) event.getData());
		pushAppEvent(AppEvts.ON_FREEZE_COLUMN, index);
	}

	@Listen("onUndo=#mainMenubar")
	public void onUndo(ForwardEvent event) {
		pushAppEvent(AppEvts.ON_UNDO);
	}

	@Listen("onRedo=#mainMenubar")
	public void onRedo(ForwardEvent event) {
		pushAppEvent(AppEvts.ON_REDO);
	}

	
	
	
	@Listen("onClick=#saveDb")
	public void saveDb() throws Exception {

		System.out.println("DB Process Started.....");

		System.out.println("MyFile Name" + bookname);
		MyProperties prop = new MyProperties();
		String driver = "org.postgresql.Driver";
		System.out.println(driver);
		String url = prop.url;
		System.out.println(url);
		String username = prop.username;
		System.out.println(username);
		String password = prop.password;
		System.out.println(password);
		String myInput_file = prop.myInput.concat(bookname);
		System.out.println("my input file :" + myInput_file);

		Connection conn = null;
		PreparedStatement sql_statement = null;
		String active_boq_file_Id = null;
		String copfilequery = null;
		String projectid = null;
		Statement sql_deleteall = null;
		Statement getrowID = null;
		Statement statement = null;
		Statement windowrefstatement = null;
		ResultSet rs = null;
		String importtype = null;
		String update_contractamt = null;
		String contractID = null;
		ResultSet value_returned = null;
		Statement sql_copstatement = null;
		System.out.println("MyFile Name :" + myInput_file);
		String window_reference = null;
		String windowId = null;
		String order_id = null;
		double product_name = 0;
		String filetype = null;
		String scenarioType = null;
		int i = 0;
		try {

			Class.forName(driver);
			conn = DriverManager.getConnection(url, username, password);
			FileInputStream myInput = new FileInputStream(myInput_file);
			XSSFWorkbook myWorkBook_contract_amt = new XSSFWorkbook(myInput);
			//Messagebox.show("Processed Succesfully : " + bookname);

			/*active_boq_file_Id = "select coalesce(m_product.name,'-'),coalesce(c_order.c_order_id,'-'),dtpo_pocopfile.filetype from c_order	
			 * RIGHT join dtpo_pocopfile on dtpo_pocopfile.c_order_id=c_ORDER.c_order_id LEFT JOIN C_ORDERLINE ON C_ORDERLINE.C_ORDER_ID=C_ORDER.C_ORDER_ID 
			 * LEFT join m_product on m_product.m_product_id=c_orderline.m_product_id	where C_ORDER.c_order_id=(select c_order_id from dtpo_pocopfile where filename='"
					+ bookname + "') and filename = '"+bookname+"'; ";*/
			
			
			active_boq_file_Id = "select c_order.c_order_id,dtpo_pocopfile.filetype," +
					"c_order.em_dtpo_orderscenario from c_order	RIGHT join dtpo_pocopfile " +
					"on dtpo_pocopfile.c_order_id=c_ORDER.c_order_id " +
					"where C_ORDER.c_order_id=(select c_order_id from dtpo_pocopfile" +
					" where filename='"+ bookname +"') and filename='"+ bookname + "'";
			
			System.out.println("sql query" + active_boq_file_Id);
			statement = conn.createStatement();
			rs = statement.executeQuery(active_boq_file_Id);
			System.out.println("active file query executed");

			while (rs.next() == true) {
				System.out.println("while condition checking");
				filetype = rs.getString(2);
				order_id = rs.getString(1);
				scenarioType = rs.getString(3);
				System.out.println("filetype "+filetype);
				System.out.println("scenarioType :"+scenarioType);
				double Rate = 0;
				//product_name = rs.getString(1);
				if (filetype.contentEquals("COP")) {
					sql_copstatement = conn.createStatement();
					ResultSet recoveryAdvance = null;  
					XSSFSheet mySheet_contractvalue = myWorkBook_contract_amt.getSheetAt(1);
					XSSFRow myrow = mySheet_contractvalue.getRow(2);
					XSSFCell mycell = myrow.getCell(3);
					XSSFRow copdocumentrow = mySheet_contractvalue.getRow(3);
					XSSFCell copdocumentcell = copdocumentrow.getCell(3);
					String copdocno = copdocumentcell.getStringCellValue();
					XSSFCell conversion= copdocumentrow.getCell(6);
					Rate = conversion.getNumericCellValue();
					if(Rate==0 || Rate==0.00)
						Rate = 1;
					//System.out.println("Lets contract id" + order_id);
					//System.out.println("Lets row_info" + row_info);
					System.out.println("Lets copdocument_no" + copdocno);
					//System.out.println("Lets product_name" + product_name);
					String copItemCode = null;
					String copItemDesc = null;
					double orderLineQTY = 0;
					double orderLineAMT = 0;
					double copBasic = 0;
					double copPercentage = 0;
					double copQuantity = 0;
					double copTotal = 0;
					double copExcise = 0;
					double copSubTotal = 0;
					double copInsurance = 0;
					double copFreight = 0;
					double copPackage = 0;
					double copOthers = 0;
					double copVAT_CST = 0;
					double copLBT = 0.0;
					double copService = 0.0;
					double copCess = 0.0;
					double copSupplyTotal = 0.0;
					double copInstallationTotal = 0.0;
					double copInstallationBasic = 0.0;
					double copInstallationService = 0.0;
					double copInstallationOthers = 0.0;
					double copImportTotal = 0.0;
					double copLocalSupplyBasic=0.0;
					//double copInstallationBasic = 0.0;
					double copSupplyOthers = 0.0;
					double rev_1strow = 0.0;
					double rev_2ndrow = 0.0;
					double rev_3rdrow = 0.0;
					double rev_4throw = 0.0;
					double rev_5throw = 0.0;
					double rev_6throw = 0.0;
					double rev_7throw = 0.0;
					double rev_8throw = 0.0;
					double rev_9throw = 0.0;
					double rev_10throw = 0.0;
					double adv_1strow = 0.0;
					double adv_2ndrow = 0.0;
					double adv_3rdrow = 0.0;
					double adv_4throw = 0.0;
					double adv_5throw = 0.0;
					double adv_6throw = 0.0;
					double TotalCop = 0.0;
					double Cop=0;
					
					//------------------------
					XSSFSheet copFaceSheet = myWorkBook_contract_amt.getSheetAt(2);
					XSSFCell copFaceSheetCell = null;
					
					copFaceSheetCell = copFaceSheet.getRow(23).getCell(6);
					if(!copFaceSheetCell.getRawValue().isEmpty()){
					rev_1strow = Double.parseDouble(copFaceSheetCell.getRawValue());
					System.out.println("Recovery row no 24 :"+rev_1strow);
					}
					copFaceSheetCell = copFaceSheet.getRow(24).getCell(6);
					if(!copFaceSheetCell.getRawValue().isEmpty()){
				    rev_2ndrow = Double.parseDouble(copFaceSheetCell.getRawValue());
				    System.out.println("Recovery row no 25 :"+rev_2ndrow);
					}
				    copFaceSheetCell = copFaceSheet.getRow(25).getCell(6);
				    if(!copFaceSheetCell.getRawValue().isEmpty()){
				    rev_3rdrow = Double.parseDouble(copFaceSheetCell.getRawValue());
				    System.out.println("Recovery row no 26 :"+rev_3rdrow);
				    }
				    copFaceSheetCell = copFaceSheet.getRow(26).getCell(6);
				    if(!copFaceSheetCell.getRawValue().isEmpty()){
				    rev_4throw = Double.parseDouble(copFaceSheetCell.getRawValue());
				    System.out.println("Recovery row no 26 :"+rev_4throw);
				    }
				    copFaceSheetCell = copFaceSheet.getRow(27).getCell(6);
				    if(!copFaceSheetCell.getRawValue().isEmpty()){
				    rev_5throw = Double.parseDouble(copFaceSheetCell.getRawValue());
				    System.out.println("Recovery row no 28 :"+rev_5throw);
				    }
				    copFaceSheetCell = copFaceSheet.getRow(28).getCell(6);
				    if(!copFaceSheetCell.getRawValue().isEmpty()){
				    rev_6throw = Double.parseDouble(copFaceSheetCell.getRawValue());
				    System.out.println("Recovery row no 29 :"+rev_6throw);
				    }
				    copFaceSheetCell = copFaceSheet.getRow(29).getCell(6);
				    if(!copFaceSheetCell.getRawValue().isEmpty()){
				    rev_7throw = Double.parseDouble(copFaceSheetCell.getRawValue());
				    System.out.println("Recovery row no 30 :"+rev_7throw);
				    }
				    copFaceSheetCell = copFaceSheet.getRow(30).getCell(6);
				    if(!copFaceSheetCell.getRawValue().isEmpty()){
				    rev_8throw = Double.parseDouble(copFaceSheetCell.getRawValue());
				    System.out.println("Recovery row no 31 :"+rev_8throw);
				    }
				    copFaceSheetCell = copFaceSheet.getRow(31).getCell(6);
				    if(!copFaceSheetCell.getRawValue().isEmpty()){
				    rev_9throw = Double.parseDouble(copFaceSheetCell.getRawValue());
				    System.out.println("Recovery row no 32 :"+rev_9throw);
				    }
				    copFaceSheetCell = copFaceSheet.getRow(32).getCell(6);
				    if(!copFaceSheetCell.getRawValue().isEmpty()){
				    rev_10throw = Double.parseDouble(copFaceSheetCell.getRawValue());
					System.out.println("Recovery row no 33 :"+rev_10throw);
				    }
					//advance static columns in excel
					copFaceSheetCell = copFaceSheet.getRow(35).getCell(6);
					if(!copFaceSheetCell.getRawValue().isEmpty()){
				    adv_1strow = Double.parseDouble(copFaceSheetCell.getRawValue());
					System.out.println("adv_1strow :"+adv_1strow);
					}
					copFaceSheetCell = copFaceSheet.getRow(36).getCell(6);
					if(!copFaceSheetCell.getRawValue().isEmpty()){
				    adv_2ndrow = Double.parseDouble(copFaceSheetCell.getRawValue());
					System.out.println("adv_2ndrow :"+adv_2ndrow);
					}
					copFaceSheetCell = copFaceSheet.getRow(37).getCell(6);
					if(!copFaceSheetCell.getRawValue().isEmpty()){
				    adv_3rdrow = Double.parseDouble(copFaceSheetCell.getRawValue());
					System.out.println("adv_3rdrow :"+adv_3rdrow);
					}
					copFaceSheetCell = copFaceSheet.getRow(38).getCell(6);
					if(!copFaceSheetCell.getRawValue().isEmpty()){
				    adv_4throw = Double.parseDouble(copFaceSheetCell.getRawValue());
					System.out.println("adv_4throw :"+adv_4throw);
					}
					copFaceSheetCell = copFaceSheet.getRow(39).getCell(6);
					if(!copFaceSheetCell.getRawValue().isEmpty()){
				    adv_5throw = Double.parseDouble(copFaceSheetCell.getRawValue());
					System.out.println("adv_5throw :"+adv_5throw);
					}
					copFaceSheetCell = copFaceSheet.getRow(40).getCell(6);
					if(!copFaceSheetCell.getRawValue().isEmpty()){
				    adv_6throw = Double.parseDouble(copFaceSheetCell.getRawValue());
					System.out.println("adv_6throw :"+adv_6throw);
					}
					copFaceSheetCell = copFaceSheet.getRow(42).getCell(6);
					System.out.println("copFaceSheetCell :"+copFaceSheetCell);
				    Cop = Double.parseDouble(copFaceSheetCell.getRawValue());
				    System.out.println("COP :"+Cop);
				    TotalCop =Cop*Rate;
					System.out.println("TotalCop :"+TotalCop);
					recoveryAdvance = sql_copstatement.executeQuery("select dtzk_update_cnvrsnpocopfile("+rev_1strow+","+rev_2ndrow+","+rev_3rdrow+","+rev_4throw+"," +
							""+rev_5throw+","+rev_6throw+","+rev_7throw+","+rev_8throw+","+rev_9throw+","+rev_10throw+"," +
									""+adv_1strow+","+adv_2ndrow+","+adv_3rdrow+","+adv_4throw+","+adv_5throw+","+adv_6throw+",'"+bookname+"','"+order_id+"',"+TotalCop+","+Cop+")from dual;");
					recoveryAdvance.next();
					System.out.println("Update ResultSet value" + recoveryAdvance.getString(1));
					/*String sqlStatement = "update dtpo_pocopile set rec_1stcolumn=? where filename=?";
					PreparedStatement updateQuery  = conn.prepareStatement(sqlStatement);
					updateQuery.setFloat(1, rev_1strow);
					updateQuery.setString(2, bookname);*/
					
					
					//-------------------------------------
					Iterator rowIter = mySheet_contractvalue.rowIterator();
					if (scenarioType.contentEquals("TYPE1")) {
						System.out.println("scenario one");
						//recovery static columns in excel
					
					while (rowIter.hasNext()) {
						XSSFRow myRow = (XSSFRow) rowIter.next();
						if (myRow.getRowNum() > 6) {
						// HSSFRow myRow = (HSSFRow) rowIter.next();
						Iterator cellIter = myRow.cellIterator();
						while (cellIter.hasNext()) {
							XSSFCell myCell = (XSSFCell) cellIter.next();
							
							 if (myCell.getColumnIndex() == 1) {		
							    copItemCode = myCell.getStringCellValue();
								System.out.println("item_code :"+copItemCode); 
                             }else if (myCell.getColumnIndex() == 2) {
								copItemDesc = myCell.getStringCellValue();
									System.out.println("copItemDesc :"+copItemDesc);
							 }else if (myCell.getColumnIndex() == 13) {
								orderLineQTY = myCell.getNumericCellValue();
									System.out.println("orderLineQTY :"+orderLineQTY);
							 }else if (myCell.getColumnIndex() == 21) {
								 orderLineAMT = myCell.getNumericCellValue();
									System.out.println("orderLineAMT :"+orderLineAMT);
							 }else if (myCell.getColumnIndex() == 23) {
								copPercentage = myCell.getNumericCellValue();
									System.out.println("copPercentage :"+copPercentage);
							 }else if (myCell.getColumnIndex() == 24) {
								copQuantity = myCell.getNumericCellValue();
									System.out.println("copQuantity :"+copQuantity);
							 }else if (myCell.getColumnIndex() == 25) {
								 //String test = myCell.toString();
								 System.out.println("test :"+myCell.getCellFormula());
								 System.out.println("test :"+myCell.getRawValue());
									copBasic = Double.parseDouble(myCell.getRawValue());
										System.out.println("copBasic :"+copBasic);
							 }else if (myCell.getColumnIndex() == 26) {
									copInsurance = Double.parseDouble(myCell.getRawValue());
									System.out.println("copInsurance :"+copInsurance);
						     }else if (myCell.getColumnIndex() == 27) {
									copFreight = Double.parseDouble(myCell.getRawValue());
									System.out.println("copFreight :"+copFreight);
						     }else if (myCell.getColumnIndex() == 28) {
									copPackage = Double.parseDouble(myCell.getRawValue());
									System.out.println("copPackage :"+copPackage);
						    }else if (myCell.getColumnIndex() == 29) {
								copOthers = Double.parseDouble(myCell.getRawValue());
								System.out.println("copOthers :"+copOthers);
					        }else if (myCell.getColumnIndex() == 30) {
								copTotal = Double.parseDouble(myCell.getRawValue());
									System.out.println("copTotal :"+copTotal);
							 }


							}
						System.out.println("COP DOCNO :" + copdocno + " Item Code :"
								+ copItemCode + "COP Item Desc :" + copItemDesc + "COP Basic :" + copBasic + "COP Percentage :" + copPercentage
								+ "COP Quantity :" + copQuantity + "COP Total :" + copTotal +
								"EXCel Name :"+bookname+"order_id :" + order_id+"copInsurance :"+copInsurance+
								"copFreight :"+copFreight+"copPackage :"+copPackage+"copOthers :"+copOthers);	

						/*value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_itemcop("+ row_info +",'" + copdocno+ "','"
								+ copItemCode + "','"+copItemDesc+"',"+copBasic+"," + copPercentage + "," + copQuantity + "," + copTotal + ",'"+bookname + "','" + order_id+ "') from dual;");*/
						
						value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_conversionRate('"+copdocno+"'," +
								"'"+copItemCode+"','"+copItemDesc+"'," + copPercentage + "," + copQuantity + "," +
										""+copBasic+","+copInsurance+","+copFreight+","+copPackage+","+copOthers+"," +
												""+copTotal+",'"+bookname + "','" + order_id+ "',"+Rate+","+orderLineQTY+","+orderLineAMT+") from dual;");
						value_returned.next();
						System.out.println("ResultSet value" + value_returned.getString(1));
						
						}
					}
					
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}	
					/*	value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_pocop('" + order_id + "'," + copqty+ ","
								+ copamt + ","+coppercent+","+copbasic+"," + product_name + ",'" + copdocno + "'," + row_info + ",'" +damagedqty + "','" + bookname
								+ "',"+product_name+") from dual;");
						value_returned.next();
						System.out.println("ResultSet value" + value_returned.getString(1));
						//	Messagebox.show("Imported Successfully");
						}

					}
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}*/
				}if (scenarioType.contentEquals("TYPE2")) {
					while (rowIter.hasNext()) {
						XSSFRow myRow = (XSSFRow) rowIter.next();
						if (myRow.getRowNum() > 6) {
						// HSSFRow myRow = (HSSFRow) rowIter.next();
						Iterator cellIter = myRow.cellIterator();
						while (cellIter.hasNext()) {
							XSSFCell myCell = (XSSFCell) cellIter.next();
							
							 if (myCell.getColumnIndex() == 1) {		
								copItemCode = myCell.getStringCellValue();
                                System.out.println("item_code :"+copItemCode);
						     }else if (myCell.getColumnIndex() == 2) {
								copItemDesc = myCell.getStringCellValue();
									System.out.println("copItemDesc :"+copItemDesc);
							 }else if (myCell.getColumnIndex() == 13) {
								orderLineQTY = myCell.getNumericCellValue();
									System.out.println("orderLineQTY :"+orderLineQTY);
							 }else if (myCell.getColumnIndex() == 25) {
								 orderLineAMT = myCell.getNumericCellValue();
									System.out.println("orderLineAMT :"+orderLineAMT);
							 }else if (myCell.getColumnIndex() == 27) {
								copPercentage = myCell.getNumericCellValue();
									System.out.println("copPercentage :"+copPercentage);
							 }else if (myCell.getColumnIndex() == 28) {
								copQuantity = myCell.getNumericCellValue();
									System.out.println("copQuantity :"+copQuantity);
							 }else if (myCell.getColumnIndex() == 29) {
								copBasic = Double.parseDouble(myCell.getRawValue());
									System.out.println("copBasic :"+copBasic);
							 }else if (myCell.getColumnIndex() == 30) {
								copExcise = Double.parseDouble(myCell.getRawValue());
									System.out.println("copExcise :"+copExcise);
							 }else if (myCell.getColumnIndex() == 31) {
								copSubTotal = Double.parseDouble(myCell.getRawValue());
									System.out.println("copSubTotal :"+copSubTotal);
							 }else if (myCell.getColumnIndex() == 32) {
								copVAT_CST = Double.parseDouble(myCell.getRawValue());
									System.out.println("copVAT_CST :"+copVAT_CST);
							 }else if (myCell.getColumnIndex() == 33) {
								copLBT = Double.parseDouble(myCell.getRawValue());
									System.out.println("copLBT :"+copLBT);
							 }else if (myCell.getColumnIndex() == 34) {
								copInsurance =Double.parseDouble(myCell.getRawValue());
									System.out.println("copInsurance :"+copInsurance);
							 }else if (myCell.getColumnIndex() == 35) {
								copFreight = Double.parseDouble(myCell.getRawValue());
									System.out.println("copFreight :"+copFreight);
							 }else if (myCell.getColumnIndex() == 36) {
								 copPackage = Double.parseDouble(myCell.getRawValue());
									System.out.println("copPackage :"+copPackage);
							 }else if (myCell.getColumnIndex() == 37) {
								copOthers = Double.parseDouble(myCell.getRawValue());
									System.out.println("copOthers :"+copOthers);
							 }else if (myCell.getColumnIndex() == 38) {
								copTotal = Double.parseDouble(myCell.getRawValue());
									System.out.println("copTotal :"+copTotal);
							 }
						}
						System.out.println("COP DOCNO :" + copdocno + " Item Code :"
								+ copItemCode + "COP Item Desc :" + copItemDesc + "COP Basic :" + copBasic + "COP Percentage :" + copPercentage
								+ "COP Quantity :" + copQuantity + "COP Total :" + copTotal + "EXCel Name :"+bookname+"order_id :" + order_id);	

					/*	value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_itemcop("+ row_info +",'" + copdocno+ "','"
								+ copItemCode + "','"+copItemDesc+"',"+copBasic+"," + copPercentage + "," + copQuantity + "," + copTotal + ",'"+bookname + "','" + order_id+ "') from dual;");*/
						value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_s2itemcop('" + copdocno+ "','"+ copItemCode + "','"+copItemDesc+"'," +
								"" + copPercentage + ","+copQuantity+","+copBasic+","+copExcise+","+copSubTotal+","+copVAT_CST+","+copLBT+"," +
										""+copInsurance+","+copFreight+","+copPackage+","+copOthers+","+copTotal+",'"+bookname + "','" + order_id+ "',"+orderLineQTY+","+orderLineAMT+") from dual;");		
						
						
						value_returned.next();
						System.out.println("ResultSet value" + value_returned.getString(1));
				      }	
					}
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}	
				}
				if (scenarioType.contentEquals("TYPE3")) {
					while (rowIter.hasNext()) {
						XSSFRow myRow = (XSSFRow) rowIter.next();
						if (myRow.getRowNum() > 6) {
						// HSSFRow myRow = (HSSFRow) rowIter.next();
						Iterator cellIter = myRow.cellIterator();
						while (cellIter.hasNext()) {
							XSSFCell myCell = (XSSFCell) cellIter.next();
							     if (myCell.getColumnIndex() == 1) {		
									copItemCode = myCell.getStringCellValue();
	                                System.out.println("item_code :"+copItemCode);
							     }else if (myCell.getColumnIndex() == 2) {
									copItemDesc = myCell.getStringCellValue();
									System.out.println("copItemDesc :"+copItemDesc);
								 }else if (myCell.getColumnIndex() == 13) {
										orderLineQTY = myCell.getNumericCellValue();
										System.out.println("orderLineQTY :"+orderLineQTY);
								 }else if (myCell.getColumnIndex() == 28) {
									 orderLineAMT = myCell.getNumericCellValue();
										System.out.println("orderLineAMT :"+orderLineAMT);
								 }else if (myCell.getColumnIndex() == 30) {
										copPercentage = myCell.getNumericCellValue();
										System.out.println("copPercentage :"+copPercentage);
								 }else if (myCell.getColumnIndex() == 31) {
									copQuantity = myCell.getNumericCellValue();
										System.out.println("copQuantity :"+copQuantity);
								 }else if (myCell.getColumnIndex() == 32) {
										copBasic = Double.parseDouble(myCell.getRawValue());
										System.out.println("copBasic :"+copBasic);
								 }else if (myCell.getColumnIndex() == 33) {
									copExcise = Double.parseDouble(myCell.getRawValue());
										System.out.println("copExcise :"+copExcise);
								 }else if (myCell.getColumnIndex() == 34) {
									copSubTotal = Double.parseDouble(myCell.getRawValue());
										System.out.println("copSubTotal :"+copSubTotal);
								 }else if (myCell.getColumnIndex() == 35) {
									copVAT_CST = Double.parseDouble(myCell.getRawValue());
										System.out.println("copVAT_CST :"+copVAT_CST);
								 }else if (myCell.getColumnIndex() == 36) {
										copOthers = Double.parseDouble(myCell.getRawValue());
										System.out.println("copOthers :"+copOthers);
								 }else if (myCell.getColumnIndex() == 37) {
										copSupplyTotal = Double.parseDouble(myCell.getRawValue());
										System.out.println("copSupplyTotal :"+copSupplyTotal);
								 }else if (myCell.getColumnIndex() == 38) {
									 copInstallationBasic = Double.parseDouble(myCell.getRawValue());
										System.out.println("copInstallationBasic :"+copInstallationBasic);
								 }else if (myCell.getColumnIndex() == 39) {
									 copInstallationService = Double.parseDouble(myCell.getRawValue());
										System.out.println("copInstallationService :"+copInstallationService);
								 }else if (myCell.getColumnIndex() == 40) {
									 copInstallationOthers = Double.parseDouble(myCell.getRawValue());
										System.out.println("copInstallationOthers :"+copInstallationOthers);
								 }else if (myCell.getColumnIndex() == 41) {
									 copInstallationTotal = Double.parseDouble(myCell.getRawValue());
										System.out.println("copInstallationTotal :"+copInstallationTotal);
								 }else if (myCell.getColumnIndex() == 42) {
										copTotal = Double.parseDouble(myCell.getRawValue());
										System.out.println("copTotal :"+copTotal);
								 }
						}
						value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_s3itemcop('" + copdocno+ "','"+ copItemCode + "','"+copItemDesc+"'," +
								"" + copPercentage + ","+copQuantity+","+copBasic+","+copExcise+","+copSubTotal+","+copVAT_CST+","+copOthers+"," +
										""+copSupplyTotal+","+copInstallationBasic+","+copInstallationService+","+copInstallationOthers+","+copInstallationTotal+","+copTotal+",'"+bookname + "','" + order_id+ "',"+orderLineQTY+","+orderLineAMT+") from dual;");		
	
						value_returned.next();
						System.out.println("ResultSet value" + value_returned.getString(1));
						}
					}
						Integer value = value_returned.getInt(1);
						if (value == 1) {
							Messagebox.show("Items Imported Successfully :" +bookname);

						}
				   }
				if (scenarioType.contentEquals("TYPE4")) {
					while (rowIter.hasNext()) {
						XSSFRow myRow = (XSSFRow) rowIter.next();
						if (myRow.getRowNum() > 6) {
						// HSSFRow myRow = (HSSFRow) rowIter.next();
						Iterator cellIter = myRow.cellIterator();
						while (cellIter.hasNext()) {
							XSSFCell myCell = (XSSFCell) cellIter.next();
							
							 if (myCell.getColumnIndex() == 1) {		
								copItemCode = myCell.getStringCellValue();
                                System.out.println("item_code :"+copItemCode);
						     }else if (myCell.getColumnIndex() == 2) {
								copItemDesc = myCell.getStringCellValue();
									System.out.println("copItemDesc :"+copItemDesc);
							 }else if (myCell.getColumnIndex() == 13) {
								orderLineQTY = myCell.getNumericCellValue();
									System.out.println("orderLineQTY :"+orderLineQTY);
							 }else if (myCell.getColumnIndex() == 21) {
								 orderLineAMT = myCell.getNumericCellValue();
									System.out.println("orderLineAMT :"+orderLineAMT);
							 }else if (myCell.getColumnIndex() == 23) {
								 copPercentage = myCell.getNumericCellValue();
									System.out.println("copPercentage :"+copPercentage);
							 }else if (myCell.getColumnIndex() == 24) {
								 copQuantity = myCell.getNumericCellValue();
									System.out.println("copQuantity :"+copQuantity);
							 }else if (myCell.getColumnIndex() == 25) {
								 copBasic = Double.parseDouble(myCell.getRawValue());
									System.out.println("copBasic :"+copBasic);
							 }else if (myCell.getColumnIndex() == 26) {
								 copVAT_CST = Double.parseDouble(myCell.getRawValue());
									System.out.println("copVAT_CST :"+copVAT_CST);
							 }else if (myCell.getColumnIndex() == 27) {
								 copService = Double.parseDouble(myCell.getRawValue());
									System.out.println("copService :"+copService);
							 }else if (myCell.getColumnIndex() == 28) {
								 copCess = Double.parseDouble(myCell.getRawValue());
									System.out.println("copCess :"+copCess);
							 }else if (myCell.getColumnIndex() == 29) {
								 copOthers = Double.parseDouble(myCell.getRawValue());
									System.out.println("copOthers :"+copOthers);
							 }else if (myCell.getColumnIndex() == 30) {
								copTotal = Double.parseDouble(myCell.getRawValue());
									System.out.println("copTotal :"+copTotal);
							 }
						}
						System.out.println("COP DOCNO :" + copdocno + " Item Code :"
								+ copItemCode + "COP Item Desc :" + copItemDesc + "COP Basic :" + copBasic + "COP Percentage :" + copPercentage
								+ "COP Quantity :" + copQuantity + "COP Total :" + copTotal + "EXCel Name :"+bookname+"order_id :" + order_id+"orderLineQTY :"+orderLineQTY+"orderLineAMT :"+orderLineAMT);	

						value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_s4itemcop('"+copdocno+ "','"+copItemCode + "','"+copItemDesc+"'," + copPercentage + "," +
								"" + copQuantity + ","+copBasic+","+copVAT_CST+","+copService+","+copCess+","+copOthers+"," + copTotal + ",'"+bookname + "','" + order_id+ "',"+orderLineQTY+","+orderLineAMT+") from dual;");
						
						
						value_returned.next();
						System.out.println("ResultSet value" + value_returned.getString(1));
				      }	
					}
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}	
				} 
				if (scenarioType.contentEquals("TYPE5")) {
					while (rowIter.hasNext()) {
						XSSFRow myRow = (XSSFRow) rowIter.next();
						if (myRow.getRowNum() > 6) {
						// HSSFRow myRow = (HSSFRow) rowIter.next();
						Iterator cellIter = myRow.cellIterator();
						while (cellIter.hasNext()) {
							XSSFCell myCell = (XSSFCell) cellIter.next();
							
							 if (myCell.getColumnIndex() == 1) {		
								copItemCode = myCell.getStringCellValue();
                                System.out.println("item_code :"+copItemCode);
						     }else if (myCell.getColumnIndex() == 2) {
								copItemDesc = myCell.getStringCellValue();
									System.out.println("copItemDesc :"+copItemDesc);
							 }else if (myCell.getColumnIndex() == 13) {
								orderLineQTY = myCell.getNumericCellValue();
									System.out.println("orderLineQTY :"+orderLineQTY);
							 }else if (myCell.getColumnIndex() == 19) {
								 orderLineAMT = myCell.getNumericCellValue();
									System.out.println("orderLineAMT :"+orderLineAMT);
							 }else if (myCell.getColumnIndex() == 21) {
								 copPercentage = myCell.getNumericCellValue();
									System.out.println("copPercentage :"+copPercentage);
							 }else if (myCell.getColumnIndex() == 22) {
								 copQuantity = Double.parseDouble(myCell.getRawValue());
									System.out.println("copQuantity :"+copQuantity);
							 }else if (myCell.getColumnIndex() == 23) {
								copBasic = Double.parseDouble(myCell.getRawValue());
									System.out.println("copBasic :"+copBasic);
							 }else if (myCell.getColumnIndex() == 24) {
								 copVAT_CST = Double.parseDouble(myCell.getRawValue());
									System.out.println("copVAT_CST :"+copVAT_CST);
							 }else if (myCell.getColumnIndex() == 25) {
								copOthers =Double.parseDouble(myCell.getRawValue());
									System.out.println("copOthers :"+copOthers);
							 }else if (myCell.getColumnIndex() == 26) {
								copTotal = Double.parseDouble(myCell.getRawValue());
									System.out.println("copTotal :"+copTotal);
							 }
						}
						/*System.out.println("COP AMOUNT :" + row_info + "COP DOCNO :" + copdocno + " Item Code :"
								+ copItemCode + "COP Item Desc :" + copItemDesc + "COP Basic :" + copBasic + "COP Percentage :" + copPercentage
								+ "COP Quantity :" + copQuantity + "COP Total :" + copTotal + "EXCel Name :"+bookname+"order_id :" + order_id);	
*/
						/*value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_itemcop("+ row_info +",'" + copdocno+ "','"
								+ copItemCode + "','"+copItemDesc+"',"+copBasic+"," + copPercentage + "," + copQuantity + "," + copTotal + ",'"+bookname + "','" + order_id+ "') from dual;");
						*/
						value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_s5itemcop('"+copdocno+"','"+ copItemCode + "','"+copItemDesc+"'," +
								"" + copPercentage + "," + copQuantity + ","+copBasic+"," +copVAT_CST + "," + copOthers + "," + copTotal + ",'"+bookname + "','" + order_id+ "',"+orderLineQTY+","+orderLineAMT+") from dual;");

						value_returned.next();
						System.out.println("ResultSet value" + value_returned.getString(1));
				      }	
					}
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}	
				} 
				if (scenarioType.contentEquals("TYPE6")) {
					System.out.println("scenario one");
					//recovery static columns in excel
				
				while (rowIter.hasNext()) {
					XSSFRow myRow = (XSSFRow) rowIter.next();
					if (myRow.getRowNum() > 6) {
					// HSSFRow myRow = (HSSFRow) rowIter.next();
					Iterator cellIter = myRow.cellIterator();
					while (cellIter.hasNext()) {
						XSSFCell myCell = (XSSFCell) cellIter.next();
						if (myCell.getColumnIndex() == 1) {		
						    copItemCode = myCell.getStringCellValue();
							System.out.println("item_code :"+copItemCode); 
                         }else if (myCell.getColumnIndex() == 2) {
							copItemDesc = myCell.getStringCellValue();
								System.out.println("copItemDesc :"+copItemDesc);
						 }else if (myCell.getColumnIndex() == 13) {
								orderLineQTY = myCell.getNumericCellValue();
								System.out.println("orderLineQTY :"+orderLineQTY);
						 }else if (myCell.getColumnIndex() == 35) {
							 orderLineAMT = myCell.getNumericCellValue();
								System.out.println("orderLineAMT :"+orderLineAMT);
						 }else if (myCell.getColumnIndex() == 37) {
							copPercentage = myCell.getNumericCellValue();
								System.out.println("copPercentage :"+copPercentage);
						 }else if (myCell.getColumnIndex() == 38) {
							copQuantity = myCell.getNumericCellValue();
								System.out.println("copQuantity :"+copQuantity);
						 }else if (myCell.getColumnIndex() == 39) {
							 	copBasic = Double.parseDouble(myCell.getRawValue());
									System.out.println("copBasic :"+copBasic);
						 }else if (myCell.getColumnIndex() == 40) {
								copInsurance = Double.parseDouble(myCell.getRawValue());
								System.out.println("copInsurance :"+copInsurance);
					     }else if (myCell.getColumnIndex() == 41) {
								copFreight = Double.parseDouble(myCell.getRawValue());
								System.out.println("copFreight :"+copFreight);
					     }else if (myCell.getColumnIndex() == 42) {
								copPackage = Double.parseDouble(myCell.getRawValue());
								System.out.println("copPackage :"+copPackage);
					     }else if (myCell.getColumnIndex() == 43) {
							copOthers = Double.parseDouble(myCell.getRawValue());
							System.out.println("copOthers :"+copOthers);
				         }else if (myCell.getColumnIndex() == 44) {
							copImportTotal = Double.parseDouble(myCell.getRawValue());
								System.out.println("copImportTotal :"+copImportTotal);
						 }else if (myCell.getColumnIndex() == 45) {
							copLocalSupplyBasic = Double.parseDouble(myCell.getRawValue());
								System.out.println("copBasic :"+copBasic);
						 }else if (myCell.getColumnIndex() == 46) {
							copExcise = Double.parseDouble(myCell.getRawValue());
								System.out.println("copExcise :"+copExcise);
						 }else if (myCell.getColumnIndex() == 47) {
							copSubTotal = Double.parseDouble(myCell.getRawValue());
								System.out.println("copSubTotal :"+copSubTotal);
						 }else if (myCell.getColumnIndex() == 48) {
							copVAT_CST = Double.parseDouble(myCell.getRawValue());
								System.out.println("copVAT_CST :"+copVAT_CST);
						 }else if (myCell.getColumnIndex() == 49) {
							copSupplyOthers = Double.parseDouble(myCell.getRawValue());
								System.out.println("copSupplyOthers :"+copSupplyOthers);
						 }else if (myCell.getColumnIndex() == 50) {
							copSupplyTotal = Double.parseDouble(myCell.getRawValue());
								System.out.println("copSupplyTotal :"+copSupplyTotal);
						 }else if (myCell.getColumnIndex() == 51) {
								copInstallationBasic = Double.parseDouble(myCell.getRawValue());
								System.out.println("copInstallationBasic :"+copInstallationBasic);
						 }else if (myCell.getColumnIndex() == 52) {
							 copInstallationService = Double.parseDouble(myCell.getRawValue());
								System.out.println("copInstallationService :"+copInstallationService);
						 }else if (myCell.getColumnIndex() == 53) {
							 copInstallationOthers =Double.parseDouble(myCell.getRawValue());
								System.out.println("copInstallationOthers :"+copInstallationOthers);
						 }else if (myCell.getColumnIndex() == 54) {
							copInstallationTotal = Double.parseDouble(myCell.getRawValue());
								System.out.println("copInstallationTotal :"+copInstallationTotal);
						 }else if (myCell.getColumnIndex() == 55) {
							copTotal = Double.parseDouble(myCell.getRawValue());
								System.out.println("copTotal :"+copTotal);
						 }
					}value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_s6itemcop('"+copdocno+"','"+copItemCode+"','"+copItemDesc+"',"+copPercentage+","+copQuantity+","+copBasic+"," +
							""+copInsurance+","+copFreight+","+copPackage+","+copOthers+","+copImportTotal+","+copLocalSupplyBasic+","+copExcise+"," +
									""+copSubTotal+","+copVAT_CST+","+copSupplyOthers+","+copSupplyTotal+","+copInstallationBasic+"," +
											""+copInstallationService+","+copInstallationOthers+","+copInstallationTotal+","+copTotal+",'"+bookname+"','"+order_id+"',"+orderLineQTY+","+orderLineAMT+") from dual;");

					value_returned.next();
					System.out.println("ResultSet value" + value_returned.getString(1));
					}
					}
				Integer value = value_returned.getInt(1);
				if (value == 1) {
					Messagebox.show("Items Imported Successfully :" +bookname);
				}
				}
				
				}	else if (filetype.contentEquals("LIMP")) {
					//Messagebox.show("Line Item Import");
					System.out.println("import condition while execute");
					
								
					String item_code =null ;
					String area = null;
					String item_no = null;
					String attribute_discription = null;
					String attribute_uom = null;
					double attribute_qty = 0;
					double attribute_rate = 0;
					String attribute_make = null;
					String attribute_model = null;
					String attribute_origin = null;
					String attribute_loading = null;
					double attribute_hsncode = 0;
					double attribute_insurance = 0;
					double attribute_freight = 0;
					double attribute_package = 0;
					double attribute_others = 0;
					String attribute_demand = null;
					double attribute_total = 0;
					double attribute_excise = 0;
					double attribute_excisetotal = 0;
					double attribute_vat_cst = 0;
					double attribute_lbt = 0;
					double attribute_service = 0;
					double attribute_cess = 0;
					double attribute_excisestr = 0;
					double attribute_installationothers = 0;
					double attribute_supplytotal = 0;
					
					double attribute_installbasic = 0;
					double attribute_servicetax = 0;
					double attribute_installationtotal = 0;
					double attribute_installationamount = 0;
					double attribute_supplyamount = 0;
					double attribute_importtotal = 0;
					double attribute_localbasic = 0;
					double attribute_localothers = 0;
					double attribute_localamount = 0;
					double attribute_supplyrate = 0;
					double attribute_supp_int_tot = 0;
					double attribute_int_total = 0;
					double conversionRate = 0;
					sql_copstatement = conn.createStatement();
					XSSFSheet mySheet_productimport = myWorkBook_contract_amt.getSheetAt(0);
					Iterator rowIter = mySheet_productimport.rowIterator();
                 if(scenarioType.contentEquals("TYPE1")){
					while (rowIter.hasNext()) {
						System.out.println("LIMP row iteration executed");
						XSSFRow myRow = (XSSFRow) rowIter.next();
						XSSFSheet mySheet_contractvalue = myWorkBook_contract_amt.getSheetAt(0);
						XSSFRow myrow = mySheet_contractvalue.getRow(3);
						//XSSFCell mycell = myrow.getCell(3);
						XSSFRow conversionRow = mySheet_contractvalue.getRow(3);
						XSSFCell conversionCell = conversionRow.getCell(3);
						//double row_info = mycell.getNumericCellValue();
						//System.out.println("Amount of COP" + row_info);
						conversionRate = conversionCell.getNumericCellValue(); 
						System.out.println("conversionRate :"+conversionRate);
						if (myRow.getRowNum() > 6) {
							System.out.println("LIMP row no greater than zero condition ");
							Iterator cellIter = myRow.cellIterator();
							while (cellIter.hasNext()) {
								System.out.println("cell iteration executed");
								XSSFCell myCell = (XSSFCell) cellIter.next();
								System.out.println("after read the xssfcell");

								
								if (myCell.getColumnIndex() == 1) {		
		                            item_code = myCell.getStringCellValue();
		                            System.out.println("item_code :"+item_code);
								}else if (myCell.getColumnIndex() == 2) {
									attribute_discription = myCell.getStringCellValue();
									System.out.println("attribute_discription :"+attribute_discription);
								}else if (myCell.getColumnIndex() == 3) {
                                     area = myCell.getStringCellValue();
                                     System.out.println("area :"+area);
								}else if (myCell.getColumnIndex() == 4) {
									attribute_make = myCell.getStringCellValue();
									System.out.println("attribute_make :"+attribute_make);
                                }else if (myCell.getColumnIndex() == 5) {
									attribute_model = myCell.getStringCellValue();
									System.out.println("attribute_model :"+attribute_model);

								}else if (myCell.getColumnIndex() == 6) {

									item_no = myCell.getStringCellValue();
									System.out.println("item_no :"+item_no);
								}else if (myCell.getColumnIndex() == 7) {
									attribute_hsncode = myCell.getNumericCellValue();
									System.out.println("attribute_hsncode :"+attribute_hsncode);
								}else if (myCell.getColumnIndex() == 8) {
									attribute_origin = myCell.getStringCellValue();
									System.out.println("attribute_origin :"+attribute_origin);
								}else if (myCell.getColumnIndex() == 9) {
									attribute_loading = myCell.getStringCellValue();
									System.out.println("attribute_loading :"+attribute_loading);
								}else if (myCell.getColumnIndex() == 10) {
									attribute_demand = myCell.getStringCellValue();
									System.out.println("attribute_demand :"+attribute_demand);
								}else if (myCell.getColumnIndex() == 12) {
									attribute_uom = myCell.getStringCellValue();
									System.out.println("attribute_uom :"+attribute_uom);
								}else if (myCell.getColumnIndex() == 13) {
									attribute_qty = myCell.getNumericCellValue();
									System.out.println("attribute_qty :"+attribute_qty);
								}else if (myCell.getColumnIndex() == 15) {
									attribute_rate = myCell.getNumericCellValue();
									System.out.println("attribute_rate :"+attribute_rate);
								}else if (myCell.getColumnIndex() == 16) {
									attribute_insurance = myCell.getNumericCellValue();
									System.out.println("attribute_insurance :"+attribute_insurance);
								}else if (myCell.getColumnIndex() == 17) {
									attribute_freight = myCell.getNumericCellValue();
									System.out.println("attribute_freight :"+attribute_freight);
								}else if (myCell.getColumnIndex() == 18) {
									attribute_package = myCell.getNumericCellValue();
									System.out.println("attribute_package :"+attribute_package);
								}else if (myCell.getColumnIndex() == 19) {
									attribute_others = myCell.getNumericCellValue();
                                    System.out.println("attribute_others :"+attribute_others);
								}else if (myCell.getColumnIndex() == 20) {
									attribute_importtotal = myCell.getNumericCellValue();
                                    System.out.println("attribute_importtotal :"+attribute_importtotal);
								}else if (myCell.getColumnIndex() == 21) {
									attribute_total = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_total);
								}
								
								 }		
								
							System.out.println("item_code :" + item_code + "area :" + area + "item_no :"
									+ item_no + "attribute_discription :" + attribute_discription + "attribute_qty :" + attribute_qty + "attribute_rate :" + attribute_rate
									+ "attribute_make :" + attribute_make + "attribute_model :" + attribute_model + "attribute_uom :"+attribute_uom+
									"attribute_origin :"+attribute_origin +"attribute_loading :"+attribute_loading +"order_id :" + order_id);	

							/*String sql = "select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "'";*/

							/*value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "') from dual;");*/
							value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_scenario1('" + item_code + "','" + area + "','" + item_no + "','"
									+ attribute_discription + "','" + attribute_origin + "','" + attribute_loading + "'," + attribute_qty + ",'" + attribute_uom + "'," + attribute_rate + ",'" + attribute_make + "','" + attribute_model + "','"
									+ order_id + "','"+attribute_demand+"',"+attribute_hsncode+",'"+attribute_insurance+"',"+attribute_freight+","+attribute_package+",'"+attribute_others+"',"+attribute_total+","+attribute_importtotal+","+conversionRate+") from dual;");		
							
							value_returned.next();
							System.out.println("ResultSet value" + value_returned.getString(1));
							
						}

					}
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}
				}else if(scenarioType.contentEquals("TYPE2")){
					while (rowIter.hasNext()) {
						System.out.println("TYPE2 row iteration executed");
						XSSFRow myRow = (XSSFRow) rowIter.next();

						if (myRow.getRowNum() > 6) {
							System.out.println("TYPE2 row no greater than zero condition ");
							Iterator cellIter = myRow.cellIterator();
							while (cellIter.hasNext()) {
								System.out.println("cell iteration executed");
								XSSFCell myCell = (XSSFCell) cellIter.next();
								System.out.println("after read the xssfcell");

								
								if (myCell.getColumnIndex() == 1) {		
		                                  item_code = myCell.getStringCellValue();
		                                  System.out.println("item_code :"+item_code);
								}else if (myCell.getColumnIndex() == 2) {
									attribute_discription = myCell.getStringCellValue();
									System.out.println("attribute_discription :"+attribute_discription);
								}else if (myCell.getColumnIndex() == 3) {
                                     area = myCell.getStringCellValue();
                                     System.out.println("area :"+area);
								}else if (myCell.getColumnIndex() == 4) {
									attribute_make = myCell.getStringCellValue();
									System.out.println("attribute_make :"+attribute_make);
                                }else if (myCell.getColumnIndex() == 5) {
									attribute_model = myCell.getStringCellValue();
									System.out.println("attribute_model :"+attribute_model);

								}else if (myCell.getColumnIndex() == 6) {

									item_no = myCell.getStringCellValue();
									System.out.println("item_no :"+item_no);
								}else if (myCell.getColumnIndex() == 7) {
									attribute_hsncode = myCell.getNumericCellValue();
									System.out.println("attribute_hsncode :"+attribute_hsncode);
								}else if (myCell.getColumnIndex() == 8) {
									attribute_origin = myCell.getStringCellValue();
									System.out.println("attribute_origin :"+attribute_origin);
								}else if (myCell.getColumnIndex() == 9) {
									attribute_loading = myCell.getStringCellValue();
									System.out.println("attribute_loading :"+attribute_loading);
								}else if (myCell.getColumnIndex() == 10) {
									attribute_demand = myCell.getStringCellValue();
									System.out.println("attribute_demand :"+attribute_demand);
								}else if (myCell.getColumnIndex() == 12) {
									attribute_uom = myCell.getStringCellValue();
									System.out.println("attribute_uom :"+attribute_uom);
								}else if (myCell.getColumnIndex() == 13) {
									attribute_qty = myCell.getNumericCellValue();
									System.out.println("attribute_qty :"+attribute_qty);
								}else if (myCell.getColumnIndex() == 15) {
									attribute_rate = myCell.getNumericCellValue();
									System.out.println("attribute_rate :"+attribute_rate);
								}else if (myCell.getColumnIndex() == 16) {
									attribute_excise = myCell.getNumericCellValue();
									System.out.println("attribute_excise :"+attribute_excise);
								}else if (myCell.getColumnIndex() == 17) {
									attribute_excisetotal = myCell.getNumericCellValue();
									System.out.println("attribute_excisetotal :"+attribute_excisetotal);
								}else if (myCell.getColumnIndex() == 18) {
									attribute_vat_cst = myCell.getNumericCellValue();
									System.out.println("attribute_vat_cst :"+attribute_vat_cst);
								}else if (myCell.getColumnIndex() == 19) {
									attribute_lbt = myCell.getNumericCellValue();
									System.out.println("attribute_lbt :"+attribute_lbt);
								}else if (myCell.getColumnIndex() == 20) {
									attribute_insurance = myCell.getNumericCellValue();
									System.out.println("attribute_insurance :"+attribute_insurance);
								}else if (myCell.getColumnIndex() == 21) {
									attribute_freight = myCell.getNumericCellValue();
									System.out.println("attribute_freight :"+attribute_freight);
								}else if (myCell.getColumnIndex() == 22) {
									attribute_package = myCell.getNumericCellValue();
									System.out.println("attribute_package :"+attribute_package);
								}else if (myCell.getColumnIndex() == 23) {
									attribute_others = myCell.getNumericCellValue();
                                    System.out.println("attribute_others :"+attribute_others);
								}else if (myCell.getColumnIndex() == 24) {
									attribute_supplyrate = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_supplyrate);
								}else if (myCell.getColumnIndex() == 25) {
									attribute_total = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_total);
								}
								
								 }		
								
							System.out.println("item_code :" + item_code + "area :" + area + "item_no :"
									+ item_no + "attribute_discription :" + attribute_discription + "attribute_qty :" + attribute_qty + "attribute_rate :" + attribute_rate
									+ "attribute_make :" + attribute_make + "attribute_model :" + attribute_model + "attribute_uom :"+attribute_uom+
									"attribute_origin :"+attribute_origin +"attribute_loading :"+attribute_loading +"order_id :" + order_id);	

							/*String sql = "select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "'";*/

							/*value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "') from dual;");*/
							value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_scenario2('" + item_code + "','" + area + "','" + item_no + "','"
									+ attribute_discription + "','" + attribute_origin + "','" + attribute_loading + "'," + attribute_qty + ",'" + attribute_uom + "'," + attribute_rate + ",'" + attribute_make + "','" + attribute_model + "','"
									+ order_id + "','"+attribute_demand+"',"+attribute_hsncode+",'"
									+attribute_insurance+"',"+attribute_freight+","+attribute_package+",'"+attribute_others+"',"+attribute_total+","
									+attribute_excise+","+attribute_excisetotal+","+attribute_vat_cst+","+attribute_lbt+","+attribute_supplyrate+") from dual;");		
							
							value_returned.next();
							System.out.println("ResultSet value" + value_returned.getString(1));
							
						}

					}
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}
				
				
                }else if(scenarioType.contentEquals("TYPE3")){
                	while (rowIter.hasNext()) {
						System.out.println("TYPE3 row iteration executed");
						XSSFRow myRow = (XSSFRow) rowIter.next();

						if (myRow.getRowNum() > 6) {
							System.out.println("TYPE3 row no greater than zero condition ");
							Iterator cellIter = myRow.cellIterator();
							while (cellIter.hasNext()) {
								System.out.println("cell iteration executed");
								XSSFCell myCell = (XSSFCell) cellIter.next();
								System.out.println("after read the xssfcell");

								
								if (myCell.getColumnIndex() == 1) {		
		                                  item_code = myCell.getStringCellValue();
		                                  System.out.println("item_code :"+item_code);
								}else if (myCell.getColumnIndex() == 2) {
									attribute_discription = myCell.getStringCellValue();
									System.out.println("attribute_discription :"+attribute_discription);
								}else if (myCell.getColumnIndex() == 3) {
                                     area = myCell.getStringCellValue();
                                     System.out.println("area :"+area);
								}else if (myCell.getColumnIndex() == 4) {
									attribute_make = myCell.getStringCellValue();
									System.out.println("attribute_make :"+attribute_make);
                                }else if (myCell.getColumnIndex() == 5) {
									attribute_model = myCell.getStringCellValue();
									System.out.println("attribute_model :"+attribute_model);

								}else if (myCell.getColumnIndex() == 6) {

									item_no = myCell.getStringCellValue();
									System.out.println("item_no :"+item_no);
								}else if (myCell.getColumnIndex() == 7) {
									attribute_hsncode = myCell.getNumericCellValue();
									System.out.println("attribute_hsncode :"+attribute_hsncode);
								}else if (myCell.getColumnIndex() == 8) {
									attribute_origin = myCell.getStringCellValue();
									System.out.println("attribute_origin :"+attribute_origin);
								}else if (myCell.getColumnIndex() == 9) {
									attribute_loading = myCell.getStringCellValue();
									System.out.println("attribute_loading :"+attribute_loading);
								}else if (myCell.getColumnIndex() == 10) {
									attribute_demand = myCell.getStringCellValue();
									System.out.println("attribute_demand :"+attribute_demand);
								}else if (myCell.getColumnIndex() == 12) {
									attribute_uom = myCell.getStringCellValue();
									System.out.println("attribute_uom :"+attribute_uom);
								}else if (myCell.getColumnIndex() == 13) {
									attribute_qty = myCell.getNumericCellValue();
									System.out.println("attribute_qty :"+attribute_qty);
								}else if (myCell.getColumnIndex() == 15) {
									attribute_rate = myCell.getNumericCellValue();
									System.out.println("attribute_rate :"+attribute_rate);
								}else if (myCell.getColumnIndex() == 16) {
									attribute_excisestr = myCell.getNumericCellValue();
									System.out.println("attribute_excise :"+attribute_excise);
								}else if (myCell.getColumnIndex() == 17) {
									attribute_excisetotal = myCell.getNumericCellValue();
									System.out.println("attribute_excisetotal :"+attribute_excisetotal);
								}else if (myCell.getColumnIndex() == 18) {
									attribute_vat_cst = myCell.getNumericCellValue();
									System.out.println("attribute_vat_cst :"+attribute_vat_cst);
								}else if (myCell.getColumnIndex() == 19) {
									attribute_others = myCell.getNumericCellValue();
									System.out.println("attribute_others :"+attribute_others);
								}else if (myCell.getColumnIndex() == 20) {
									attribute_supplytotal = myCell.getNumericCellValue();
									System.out.println("attribute_supplytotal :"+attribute_supplytotal);
								}else if (myCell.getColumnIndex() == 22) {
									attribute_installbasic = myCell.getNumericCellValue();
									System.out.println("attribute_installbasic :"+attribute_installbasic);
								}else if (myCell.getColumnIndex() == 23) {
									attribute_servicetax = myCell.getNumericCellValue();
									System.out.println("attribute_servicetax :"+attribute_servicetax);
								}else if (myCell.getColumnIndex() == 24) {
									attribute_installationothers = myCell.getNumericCellValue();
                                    System.out.println("attribute_installationothers :"+attribute_installationothers);
								}else if (myCell.getColumnIndex() == 25) {
									attribute_installationtotal = myCell.getNumericCellValue();
                                    System.out.println("attribute_installationtotal :"+attribute_installationtotal);
								}else if (myCell.getColumnIndex() == 26) {
									attribute_supplyamount = myCell.getNumericCellValue();
                                    System.out.println("attribute_supplyamount :"+attribute_supplyamount);
								}else if (myCell.getColumnIndex() == 27) {
									attribute_installationamount = myCell.getNumericCellValue();
                                    System.out.println("attribute_installationamount :"+attribute_installationamount);
								}else if (myCell.getColumnIndex() == 28) {
									attribute_total = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_total);
								}
								
								 }		
								
							System.out.println("item_code :" + item_code + "area :" + area + "item_no :"
									+ item_no + "attribute_discription :" + attribute_discription + "attribute_qty :" + attribute_qty + "attribute_rate :" + attribute_rate
									+ "attribute_make :" + attribute_make + "attribute_model :" + attribute_model + "attribute_uom :"+attribute_uom+
									"attribute_origin :"+attribute_origin +"attribute_loading :"+attribute_loading +"order_id :" + order_id);	

							/*String sql = "select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "'";*/

							/*value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "') from dual;");*/
							value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_scenario3('" + item_code + "','" + area + "','" + item_no + "','"
									+ attribute_discription + "','" + attribute_origin + "','" + attribute_loading + "'," + attribute_qty + ",'" + attribute_uom + "'," + attribute_rate + ",'" + attribute_make + "','" + attribute_model + "','"
									+ order_id + "','"+attribute_demand+"',"+attribute_hsncode+","+attribute_excisestr+","+attribute_excisetotal+","+attribute_vat_cst+","+attribute_others+","+attribute_supplytotal+","+attribute_installbasic+","+attribute_servicetax+","+attribute_installationothers+","+attribute_installationtotal+","+attribute_supplyamount+","+attribute_installationamount+","+attribute_total+") from dual;");		
							
							value_returned.next();
							System.out.println("ResultSet value" + value_returned.getString(1));
							
						}

					}
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}
                }else if(scenarioType.contentEquals("TYPE4")){
                	while (rowIter.hasNext()) {
						System.out.println("LIMP row iteration executed");
						XSSFRow myRow = (XSSFRow) rowIter.next();

						if (myRow.getRowNum() > 6) {
							System.out.println("LIMP row no greater than zero condition ");
							Iterator cellIter = myRow.cellIterator();
							while (cellIter.hasNext()) {
								System.out.println("cell iteration executed");
								XSSFCell myCell = (XSSFCell) cellIter.next();
								System.out.println("after read the xssfcell");

								
								if (myCell.getColumnIndex() == 1) {		
		                                  item_code = myCell.getStringCellValue();
		                                  System.out.println("item_code :"+item_code);
								}else if (myCell.getColumnIndex() == 2) {
									attribute_discription = myCell.getStringCellValue();
									System.out.println("attribute_discription :"+attribute_discription);
								}else if (myCell.getColumnIndex() == 3) {
                                     area = myCell.getStringCellValue();
                                     System.out.println("area :"+area);
								}else if (myCell.getColumnIndex() == 4) {
									attribute_make = myCell.getStringCellValue();
									System.out.println("attribute_make :"+attribute_make);
                                }else if (myCell.getColumnIndex() == 5) {
									attribute_model = myCell.getStringCellValue();
									System.out.println("attribute_model :"+attribute_model);

								}else if (myCell.getColumnIndex() == 6) {

									item_no = myCell.getStringCellValue();
									System.out.println("item_no :"+item_no);
								}else if (myCell.getColumnIndex() == 7) {
									attribute_hsncode = myCell.getNumericCellValue();
									System.out.println("attribute_hsncode :"+attribute_hsncode);
								}else if (myCell.getColumnIndex() == 8) {
									attribute_origin = myCell.getStringCellValue();
									System.out.println("attribute_origin :"+attribute_origin);
								}else if (myCell.getColumnIndex() == 9) {
									attribute_loading = myCell.getStringCellValue();
									System.out.println("attribute_loading :"+attribute_loading);
								}else if (myCell.getColumnIndex() == 10) {
									attribute_demand = myCell.getStringCellValue();
									System.out.println("attribute_demand :"+attribute_demand);
								}else if (myCell.getColumnIndex() == 12) {
									attribute_uom = myCell.getStringCellValue();
									System.out.println("attribute_uom :"+attribute_uom);
								}else if (myCell.getColumnIndex() == 13) {
									attribute_qty = myCell.getNumericCellValue();
									System.out.println("attribute_qty :"+attribute_qty);
								}else if (myCell.getColumnIndex() == 15) {
									attribute_rate = myCell.getNumericCellValue();
									System.out.println("attribute_rate :"+attribute_rate);
								}else if (myCell.getColumnIndex() == 16) {
									attribute_vat_cst = myCell.getNumericCellValue();
									System.out.println("attribute_vat_cst :"+attribute_vat_cst);
								}else if (myCell.getColumnIndex() == 17) {
									attribute_service = myCell.getNumericCellValue();
									System.out.println("attribute_service :"+attribute_service);
								}else if (myCell.getColumnIndex() == 18) {
									attribute_cess = myCell.getNumericCellValue();
									System.out.println("attribute_cess :"+attribute_cess);
								}else if (myCell.getColumnIndex() == 19) {
									attribute_others = myCell.getNumericCellValue();
                                    System.out.println("attribute_others :"+attribute_others);
								}else if (myCell.getColumnIndex() == 20) {
									attribute_supp_int_tot = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_supp_int_tot);
								}else if (myCell.getColumnIndex() == 21) {
									attribute_total = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_total);
								}
								
								 }		
								
							System.out.println("item_code :" + item_code + "area :" + area + "item_no :"
									+ item_no + "attribute_discription :" + attribute_discription + "attribute_qty :" + attribute_qty + "attribute_rate :" + attribute_rate
									+ "attribute_make :" + attribute_make + "attribute_model :" + attribute_model + "attribute_uom :"+attribute_uom+
									"attribute_origin :"+attribute_origin +"attribute_loading :"+attribute_loading +"order_id :" + order_id);	

							/*String sql = "select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "'";*/

							/*value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "') from dual;");*/
							value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_scenario4('" + item_code + "','" + area + "','" + item_no + "','"
									+ attribute_discription + "','" + attribute_origin + "','" + attribute_loading + "'," + attribute_qty + ",'" + attribute_uom + "'," + attribute_rate + ",'" + attribute_make + "','" + attribute_model + "','"
									+ order_id + "','"+attribute_demand+"',"+attribute_hsncode+","+attribute_vat_cst+","+attribute_service+","+attribute_cess+","+attribute_others+","+attribute_total+","+attribute_supp_int_tot+") from dual;");		
							
							value_returned.next();
							System.out.println("ResultSet value" + value_returned.getString(1));
							
						}

					}
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}
                	
                }else if(scenarioType.contentEquals("TYPE5")){
                	while (rowIter.hasNext()) {
						System.out.println("LIMP row iteration executed");
						XSSFRow myRow = (XSSFRow) rowIter.next();

						if (myRow.getRowNum() > 6) {
							System.out.println("LIMP row no greater than zero condition ");
							Iterator cellIter = myRow.cellIterator();
							while (cellIter.hasNext()) {
								System.out.println("cell iteration executed");
								XSSFCell myCell = (XSSFCell) cellIter.next();
								System.out.println("after read the xssfcell");

								
								if (myCell.getColumnIndex() == 1) {		
		                                  item_code = myCell.getStringCellValue();
		                                  System.out.println("item_code :"+item_code);
								}else if (myCell.getColumnIndex() == 2) {
									attribute_discription = myCell.getStringCellValue();
									System.out.println("attribute_discription :"+attribute_discription);
								}else if (myCell.getColumnIndex() == 3) {
                                     area = myCell.getStringCellValue();
                                     System.out.println("area :"+area);
								}else if (myCell.getColumnIndex() == 4) {
									attribute_make = myCell.getStringCellValue();
									System.out.println("attribute_make :"+attribute_make);
                                }else if (myCell.getColumnIndex() == 5) {
									attribute_model = myCell.getStringCellValue();
									System.out.println("attribute_model :"+attribute_model);

								}else if (myCell.getColumnIndex() == 6) {

									item_no = myCell.getStringCellValue();
									System.out.println("item_no :"+item_no);
								}else if (myCell.getColumnIndex() == 7) {
									attribute_hsncode = myCell.getNumericCellValue();
									System.out.println("attribute_hsncode :"+attribute_hsncode);
								}else if (myCell.getColumnIndex() == 8) {
									attribute_origin = myCell.getStringCellValue();
									System.out.println("attribute_origin :"+attribute_origin);
								}else if (myCell.getColumnIndex() == 9) {
									attribute_loading = myCell.getStringCellValue();
									System.out.println("attribute_loading :"+attribute_loading);
								}else if (myCell.getColumnIndex() == 10) {
									attribute_demand = myCell.getStringCellValue();
									System.out.println("attribute_demand :"+attribute_demand);
								}else if (myCell.getColumnIndex() == 12) {
									attribute_uom = myCell.getStringCellValue();
									System.out.println("attribute_uom :"+attribute_uom);
								}else if (myCell.getColumnIndex() == 13) {
									attribute_qty = myCell.getNumericCellValue();
									System.out.println("attribute_qty :"+attribute_qty);
								}else if (myCell.getColumnIndex() == 15) {
									attribute_rate = myCell.getNumericCellValue();
									System.out.println("attribute_rate :"+attribute_rate);
								}else if (myCell.getColumnIndex() == 16) {
									attribute_service = myCell.getNumericCellValue();
									System.out.println("attribute_service :"+attribute_service);
								}else if (myCell.getColumnIndex() == 17) {
									attribute_others = myCell.getNumericCellValue();
                                    System.out.println("attribute_others :"+attribute_others);
								}else if (myCell.getColumnIndex() == 18) {
									attribute_int_total = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_int_total);
								}else if (myCell.getColumnIndex() == 19) {
									attribute_total = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_total);
								}
								
								 }		
								
							System.out.println("item_code :" + item_code + "area :" + area + "item_no :"
									+ item_no + "attribute_discription :" + attribute_discription + "attribute_qty :" + attribute_qty + "attribute_rate :" + attribute_rate
									+ "attribute_make :" + attribute_make + "attribute_model :" + attribute_model + "attribute_uom :"+attribute_uom+
									"attribute_origin :"+attribute_origin +"attribute_loading :"+attribute_loading +"order_id :" + order_id);	

							/*String sql = "select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "'";*/

							/*value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "') from dual;");*/
							value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_scenario5('" + item_code + "','" + area + "','" + item_no + "','"
									+ attribute_discription + "','" + attribute_origin + "','" + attribute_loading + "'," + attribute_qty + ",'" + attribute_uom + "'," + attribute_rate + ",'" + attribute_make + "','" + attribute_model + "','"
									+ order_id + "','"+attribute_demand+"',"+attribute_hsncode+","+attribute_service+","+attribute_others+","+attribute_total+","+attribute_int_total+") from dual;");		
							
							value_returned.next();
							System.out.println("ResultSet value" + value_returned.getString(1));
							
						}

					}
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}
                	
                	
                }else if(scenarioType.contentEquals("TYPE6")){
                	while (rowIter.hasNext()) {
						System.out.println("LIMP row iteration executed");
						XSSFRow myRow = (XSSFRow) rowIter.next();

						if (myRow.getRowNum() > 6) {
							System.out.println("LIMP row no greater than zero condition ");
							Iterator cellIter = myRow.cellIterator();
							while (cellIter.hasNext()) {
								System.out.println("cell iteration executed");
								XSSFCell myCell = (XSSFCell) cellIter.next();
								System.out.println("after read the xssfcell");

								
								if (myCell.getColumnIndex() == 1) {		
		                                  item_code = myCell.getStringCellValue();
		                                  System.out.println("item_code :"+item_code);
								}else if (myCell.getColumnIndex() == 2) {
									attribute_discription = myCell.getStringCellValue();
									System.out.println("attribute_discription :"+attribute_discription);
								}else if (myCell.getColumnIndex() == 3) {
                                     area = myCell.getStringCellValue();
                                     System.out.println("area :"+area);
								}else if (myCell.getColumnIndex() == 4) {
									attribute_make = myCell.getStringCellValue();
									System.out.println("attribute_make :"+attribute_make);
                                }else if (myCell.getColumnIndex() == 5) {
									attribute_model = myCell.getStringCellValue();
									System.out.println("attribute_model :"+attribute_model);

								}else if (myCell.getColumnIndex() == 6) {

									item_no = myCell.getStringCellValue();
									System.out.println("item_no :"+item_no);
								}else if (myCell.getColumnIndex() == 7) {
									attribute_hsncode = myCell.getNumericCellValue();
									System.out.println("attribute_hsncode :"+attribute_hsncode);
								}else if (myCell.getColumnIndex() == 8) {
									attribute_origin = myCell.getStringCellValue();
									System.out.println("attribute_origin :"+attribute_origin);
								}else if (myCell.getColumnIndex() == 9) {
									attribute_loading = myCell.getStringCellValue();
									System.out.println("attribute_loading :"+attribute_loading);
								}else if (myCell.getColumnIndex() == 10) {
									attribute_demand = myCell.getStringCellValue();
									System.out.println("attribute_demand :"+attribute_demand);
								}else if (myCell.getColumnIndex() == 12) {
									attribute_uom = myCell.getStringCellValue();
									System.out.println("attribute_uom :"+attribute_uom);
								}else if (myCell.getColumnIndex() == 13) {
									attribute_qty = myCell.getNumericCellValue();
									System.out.println("attribute_qty :"+attribute_qty);
								}else if (myCell.getColumnIndex() == 15) {
									attribute_rate = myCell.getNumericCellValue();
									System.out.println("attribute_rate :"+attribute_rate);
								}else if (myCell.getColumnIndex() == 16) {
									attribute_insurance = myCell.getNumericCellValue();
									System.out.println("attribute_insurance :"+attribute_insurance);
								}else if (myCell.getColumnIndex() == 17) {
									attribute_freight = myCell.getNumericCellValue();
									System.out.println("attribute_freight :"+attribute_freight);
								}else if (myCell.getColumnIndex() == 18) {
									attribute_package = myCell.getNumericCellValue();
									System.out.println("attribute_package :"+attribute_package);
								}else if (myCell.getColumnIndex() == 19) {
									attribute_others = myCell.getNumericCellValue();
                                    System.out.println("attribute_others :"+attribute_others);
								}else if (myCell.getColumnIndex() == 20) {
									attribute_importtotal = myCell.getNumericCellValue();
                                    System.out.println("attribute_importtotal :"+attribute_importtotal);
								}else if (myCell.getColumnIndex() == 21) {
									attribute_localbasic = myCell.getNumericCellValue();
                                    System.out.println("attribute_localbasic :"+attribute_localbasic);
								}else if (myCell.getColumnIndex() == 22) {
									attribute_excisestr = myCell.getNumericCellValue();
                                    System.out.println("attribute_excisestr :"+attribute_excisestr);
								}else if (myCell.getColumnIndex() == 23) {
									attribute_excisetotal = myCell.getNumericCellValue();
                                    System.out.println("attribute_excisetotal :"+attribute_excisetotal);
								}else if (myCell.getColumnIndex() == 24) {
									attribute_vat_cst = myCell.getNumericCellValue();
                                    System.out.println("attribute_vat_cst :"+attribute_vat_cst);
								}else if (myCell.getColumnIndex() == 25) {
									attribute_localothers = myCell.getNumericCellValue();
                                    System.out.println("attribute_localothers :"+attribute_localothers);
								}else if (myCell.getColumnIndex() == 26) {
									attribute_supplytotal = myCell.getNumericCellValue();
                                    System.out.println("attribute_supplytotal :"+attribute_supplytotal);
								}else if (myCell.getColumnIndex() == 28) {
									attribute_installbasic = myCell.getNumericCellValue();
                                    System.out.println("attribute_installbasic :"+attribute_installbasic);
								}else if (myCell.getColumnIndex() == 29) {
									attribute_servicetax = myCell.getNumericCellValue();
                                    System.out.println("attribute_servicetax :"+attribute_servicetax);
								}else if (myCell.getColumnIndex() == 30) {
									attribute_installationothers = myCell.getNumericCellValue();
                                    System.out.println("attribute_installationothers :"+attribute_installationothers);
								}else if (myCell.getColumnIndex() == 31) {
									attribute_installationtotal = myCell.getNumericCellValue();
                                    System.out.println("attribute_installationtotal :"+attribute_installationtotal);
								}else if (myCell.getColumnIndex() == 32) {
									attribute_supplyamount = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_supplyamount);
								}else if (myCell.getColumnIndex() == 33) {
									attribute_localamount = myCell.getNumericCellValue();
									System.out.println("attribute_localamount :"+attribute_localamount);
								}else if (myCell.getColumnIndex() == 34) {
									attribute_installationamount = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_installationamount);
								}else if (myCell.getColumnIndex() == 35) {
									attribute_total = myCell.getNumericCellValue();
									System.out.println("attribute_total :"+attribute_total);
								}
								
								 }		
								
							System.out.println("item_code :" + item_code + "area :" + area + "item_no :"
									+ item_no + "attribute_discription :" + attribute_discription + "attribute_qty :" + attribute_qty + "attribute_rate :" + attribute_rate
									+ "attribute_make :" + attribute_make + "attribute_model :" + attribute_model + "attribute_uom :"+attribute_uom+
									"attribute_origin :"+attribute_origin +"attribute_loading :"+attribute_loading +"order_id :" + order_id);	

							/*String sql = "select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "'";*/

							/*value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_poline('" + product_category + "','" + product_search + "','" + item_name + "','"
									+ product_discription + "','" + item_uom + "'," + item_qty + "," + item_rate + "," + item_amount + ",'"
									+ order_id + "') from dual;");*/
							value_returned = sql_copstatement.executeQuery("select dtzk_dataimport_scenario6('" + item_code + "','" + attribute_discription + "','" + area + "','"
									+ attribute_make + "','" + attribute_model + "','" + item_no + "'," + attribute_hsncode + ",'" + attribute_origin + "','" + attribute_loading + "','" + attribute_demand + "','" + attribute_uom + "',"
									+ attribute_qty + ","+attribute_rate+","+attribute_insurance+","+attribute_freight+","+attribute_package+","+attribute_others+","+attribute_importtotal+","+attribute_localbasic+","+attribute_excisestr+","+attribute_excisetotal+","+attribute_vat_cst+","+attribute_localothers+","+attribute_supplytotal+","+attribute_installbasic+","+attribute_servicetax+","+attribute_installationothers+","+attribute_installationtotal+","+attribute_supplyamount+","+attribute_localamount+","+attribute_installationamount+","+attribute_total+",'"+order_id+"') from dual;");		
							
							value_returned.next();
							System.out.println("ResultSet value" + value_returned.getString(1));
							
						}

					}
					Integer value = value_returned.getInt(1);
					if (value == 1) {
						Messagebox.show("Items Imported Successfully :" +bookname);

					}

				}
				}
			}
		} catch (NullPointerException e) {
			e.printStackTrace();

		}

		

	}
	@Listen("onClick=#VendorDetails")
	public void VendorDetails() throws Exception {
	 
		
		System.out.println("bookname: " + bookname);
		
				System.out.println("VendorDetails button executed");
		MyProperties prop = new MyProperties();
		String driver = "org.postgresql.Driver";
		System.out.println(driver);
		String url = prop.url;
		System.out.println(url);
		String username = prop.username;
		System.out.println(username);
		String password = prop.password;
		System.out.println(password);
		String myFileName = prop.myInput.concat(bookname);
		System.out.println("my input file" + myFileName);

		Connection conn = null;
		ResultSet selectOrderData = null;
		ResultSet selectVendorDetailsData=null;
		ResultSet selectCopData=null;
		ResultSet selectCopValueData=null;
		
		Statement statement = null;
		
		String orderScenario = null;
		String orderId = null;
		
		String copCount = null;
		String documentNo = null;
		String grandTotal = null;
		Date orderDate=null;
		String bpartnerName = null;
		String address1 = null;
		String address2 = null;
		String city = null;
		String postal = null;
		String countryName = null;
		String panNo = null;
		String vatNo = null;
		String invoiceNo = null;
		//String copValue = null;
		String copNoValue =null;
		String previousCop =null;
		try{
			
			Class.forName(driver);
			conn = DriverManager.getConnection(url, username, password);
			statement = conn.createStatement();
			String selectOrderIdQuery = "select c_order_id from dtpo_pocopfile where filename='"+bookname+"'";
			System.out.println("selectOrderIdQuery: "+selectOrderIdQuery);
			selectOrderData=statement.executeQuery(selectOrderIdQuery);
 
			
			while(selectOrderData.next()==true){
				orderId = selectOrderData.getString(1);
				System.out.println("Order Id: "+orderId);
			 	}
		
			String selectVendorDetailsQuery = "select c_order.documentno,c_order.grandtotal,c_order.dateordered,c_bpartner.name,c_location.address1,c_location.address2,"
					+ "c_location.city,c_location.postal,c_country.name,c_bpartner.taxid, c_bpartner.description,c_bpartner.referenceno "
					+ "from c_order left join c_bpartner on c_order.c_bpartner_id=c_bpartner.c_bpartner_id"
					+ " left join c_bpartner_location on c_bpartner.c_bpartner_id=c_bpartner_location.c_bpartner_id"
					+ " left join c_location on c_bpartner_location.c_location_id=c_location.c_location_id"
					+ " left join c_country on c_location.c_country_id=c_country.c_country_id where c_order_id='"+orderId+"'";
			selectVendorDetailsData=statement.executeQuery(selectVendorDetailsQuery);
			
			while(selectVendorDetailsData.next()==true){
				documentNo = selectVendorDetailsData.getString(1);
				System.out.println("vendorDetails Type documentNo:"+documentNo);
				
				grandTotal = selectVendorDetailsData.getString(2);
				System.out.println("vendorDetails Type grandTotal:"+grandTotal);
				
				orderDate = selectVendorDetailsData.getDate(3);
				System.out.println("vendorDetails Type orderDate:"+orderDate);
				
				bpartnerName = selectVendorDetailsData.getString(4);
				System.out.println("vendorDetails Type bpartnerName:"+bpartnerName);
				
				address1 = selectVendorDetailsData.getString(5);
				System.out.println("vendorDetails Type address1:"+address1);
				
				address2 = selectVendorDetailsData.getString(6);
				System.out.println("vendorDetails Type address2:"+address2);
				
				city = selectVendorDetailsData.getString(7);
				System.out.println("vendorDetails Type city:"+city);
				
				postal = selectVendorDetailsData.getString(8);
				System.out.println("vendorDetails Type postal:"+postal);
				
				countryName = selectVendorDetailsData.getString(9);
				System.out.println("vendorDetails Type countryName:"+countryName);
				
				panNo = selectVendorDetailsData.getString(10);
				System.out.println("vendorDetails Type panNo:"+panNo);
				
				vatNo = selectVendorDetailsData.getString(11);
				System.out.println("vendorDetails Type vatNo:"+vatNo);
				
				invoiceNo = selectVendorDetailsData.getString(12);
				System.out.println("vendorDetails Type invoiceNo:"+invoiceNo);
		
			}
			System.out.println("test1...");
			String selectCopNoQuery="select count(*) from dtpo_pocopfile where filetype='COP' and c_order_id='"+orderId+"'";
			selectCopData=statement.executeQuery(selectCopNoQuery);
			while(selectCopData.next()==true){
				
				copCount = selectCopData.getString(1);
				System.out.println("copCount: "+copCount);
			
				int copNo = Integer.parseInt(copCount);
				System.out.println("copNumber: "+copNo);
				
				if(copNo==0){
					 copValue="COP-R001";	
				}
				else{
					/*String selectCopValueQuery="select max(substr(copno,(length(copno)-2),3)) from dtpo_pocopfile where c_order_id='"+orderId+"'";
					selectCopValueData=statement.executeQuery(selectCopValueQuery);
				
					while(selectCopValueData.next()==true){
						copNoValue = selectCopValueData.getString(1);
						System.out.println("copNoValue: "+copNoValue);
						
						int subcop = copNoValue.length();
						System.out.println("subcop: "+subcop);*/
						
						/*int finalcop = copNumber + 1;
						System.out.println("finalcop: "+finalcop);
						copValue=Integer.toString(finalcop);*/
					String selectCopValueQuery="select max(copno) from dtpo_pocopfile where c_order_id='"+orderId+"'";
				selectCopValueData=statement.executeQuery(selectCopValueQuery);			
				while(selectCopValueData.next()==true){
					previousCop = selectCopValueData.getString(1);
						System.out.println("previouscop: "+previousCop);
				}
				
						copValue="COP-R00"+copNo;
						System.out.println("copValue: "+copValue);
						
					
				
		
			
			FileInputStream myInput = new FileInputStream(myFileName);
			XSSFWorkbook myWorkBookSheets = new XSSFWorkbook(myInput);
			System.out.println("myWorkBookSheets :"+myWorkBookSheets.getNumberOfSheets());
			//System.out.println("myWorkBook_contract_amt myInput :");
			
			XSSFSheet myFirstSheet = myWorkBookSheets.getSheetAt(1);
			System.out.println("myFirstSheet: "+myFirstSheet);
			//XSSFCell myFirstShetCell = null;
            XSSFSheet mySecondSheet = myWorkBookSheets.getSheetAt(2);
			System.out.println("mySecondSheet: "+mySecondSheet);
			
		 	XSSFCell mycell = null;
			
			mycell=mySecondSheet.getRow(7).getCell(3);
			System.out.println("orderdate cell :"+mycell);
			mycell.setCellValue(documentNo+"/"+orderDate);
			
			mycell=mySecondSheet.getRow(5).getCell(5);
			System.out.println("bpname cell :"+mycell);
			mycell.setCellValue(bpartnerName);
			
			mycell=mySecondSheet.getRow(6).getCell(5);
			System.out.println("bpaddr cell :"+mycell);
			mycell.setCellValue(address1+address2+","+city+","+postal+","+countryName);
			
			mycell=mySecondSheet.getRow(7).getCell(5);
			System.out.println("panno cell :"+mycell);
			mycell.setCellValue(panNo);
			
			mycell=mySecondSheet.getRow(8).getCell(5);
			System.out.println("vatno cell :"+mycell);
			mycell.setCellValue(vatNo);
			
			mycell=mySecondSheet.getRow(9).getCell(5);
			System.out.println("invoiceno cell :"+mycell);
			mycell.setCellValue(invoiceNo);
			
			mycell=mySecondSheet.getRow(8).getCell(3);
			System.out.println("grandtotal cell :"+mycell);
			mycell.setCellValue("[$INR] "+grandTotal);
			
			mycell=myFirstSheet.getRow(3).getCell(3);
			System.out.println("copvalue 1stsheet cell :"+mycell);
			mycell.setCellValue(copValue);
			
			mycell=mySecondSheet.getRow(14).getCell(6);
			System.out.println("copvalue 2ndsheet cell :"+mycell);
			mycell.setCellValue(copValue);
			
			mycell=mySecondSheet.getRow(14).getCell(5);
			System.out.println("prevcop cell :"+mycell);
			mycell.setCellValue(previousCop);
			
			
			FileOutputStream fileOut = new FileOutputStream(myFileName);
			myWorkBookSheets.write(fileOut);
			fileOut.close();
			Messagebox.show("Vendor details Imported Successfully "+bookname);
			
			}
			}
		}catch (NullPointerException e) {
			e.printStackTrace();

		}
	}


	@Listen("onClick=#PreviousCOP")
	public void PreviousCOP() throws Exception {
		
		System.out.println("PreviousCOP button executed");
		MyProperties prop = new MyProperties();
		String driver = "org.postgresql.Driver";
		System.out.println(driver);
		String url = prop.url;
		System.out.println(url);
		String username = prop.username;
		System.out.println(username);
		String password = prop.password;
		System.out.println(password);
		String copFileName = prop.myInput.concat(bookname);
		System.out.println("my input file" + copFileName);

		Connection conn = null;/*PreparedStatement sql_statement = null;*/
		
		String selectPrevCOPData = null;
		String selectCOPBasicData = null;
		ResultSet orderID = null;
		ResultSet prevCOPData = null;
		
		
		String previousTotals = null;
		
		Statement statement = null;
		
		ResultSet rs = null;
		ResultSet scenarioType = null;
		String orderScenario = null;
		
		System.out.println("MyFile Name" + copFileName);
		
		System.out.println("bookname: " + bookname);
		
		System.out.println("VendorDetails button executed");
		/*MyProperties prop = new MyProperties();
		String driver = "org.postgresql.Driver";
		System.out.println(driver);
		String url = prop.url;
		System.out.println(url);
		String username = prop.username;
		System.out.println(username);
		String password = prop.password;
		System.out.println(password);*/
		String myFileName = prop.myInput.concat(bookname);
		System.out.println("my input file" + myFileName);

//Connection conn = null;
		ResultSet selectOrderData = null;
		ResultSet selectVendorDetailsData=null;
		ResultSet selectCopData=null;
		ResultSet selectCopValueData=null;

//Statement statement = null;

//String orderScenario = null;
		String orderId = null;

		String copCount = null;
		String documentNo = null;
		String grandTotal = null;
		Date orderDate=null;
		String bpartnerName = null;
		String address1 = null;
		String address2 = null;
		String city = null;
		String postal = null;
		String countryName = null;
		String panNo = null;
		String vatNo = null;
		String invoiceNo = null;
//String copValue = null;
		String copNoValue =null;
		String previousCop =null;
		try{
	
	Class.forName(driver);
	conn = DriverManager.getConnection(url, username, password);
	statement = conn.createStatement();
	String selectOrderIdQuery = "select c_order_id from dtpo_pocopfile where filename='"+bookname+"'";
	System.out.println("selectOrderIdQuery: "+selectOrderIdQuery);
	selectOrderData=statement.executeQuery(selectOrderIdQuery);

	
	while(selectOrderData.next()==true){
		orderId = selectOrderData.getString(1);
		System.out.println("Order Id: "+orderId);
	 	}

	String selectVendorDetailsQuery = "select c_order.documentno,c_order.grandtotal,c_order.dateordered,c_bpartner.name,c_location.address1,c_location.address2,"
			+ "c_location.city,c_location.postal,c_country.name,c_bpartner.taxid, c_bpartner.description,c_bpartner.referenceno "
			+ "from c_order left join c_bpartner on c_order.c_bpartner_id=c_bpartner.c_bpartner_id"
			+ " left join c_bpartner_location on c_bpartner.c_bpartner_id=c_bpartner_location.c_bpartner_id"
			+ " left join c_location on c_bpartner_location.c_location_id=c_location.c_location_id"
			+ " left join c_country on c_location.c_country_id=c_country.c_country_id where c_order_id='"+orderId+"'";
	selectVendorDetailsData=statement.executeQuery(selectVendorDetailsQuery);
	
	while(selectVendorDetailsData.next()==true){
		documentNo = selectVendorDetailsData.getString(1);
		System.out.println("vendorDetails Type documentNo:"+documentNo);
		
		grandTotal = selectVendorDetailsData.getString(2);
		System.out.println("vendorDetails Type grandTotal:"+grandTotal);
		
		orderDate = selectVendorDetailsData.getDate(3);
		System.out.println("vendorDetails Type orderDate:"+orderDate);
		
		bpartnerName = selectVendorDetailsData.getString(4);
		System.out.println("vendorDetails Type bpartnerName:"+bpartnerName);
		
		address1 = selectVendorDetailsData.getString(5);
		System.out.println("vendorDetails Type address1:"+address1);
		
		address2 = selectVendorDetailsData.getString(6);
		System.out.println("vendorDetails Type address2:"+address2);
		
		city = selectVendorDetailsData.getString(7);
		System.out.println("vendorDetails Type city:"+city);
		
		postal = selectVendorDetailsData.getString(8);
		System.out.println("vendorDetails Type postal:"+postal);
		
		countryName = selectVendorDetailsData.getString(9);
		System.out.println("vendorDetails Type countryName:"+countryName);
		
		panNo = selectVendorDetailsData.getString(10);
		System.out.println("vendorDetails Type panNo:"+panNo);
		
		vatNo = selectVendorDetailsData.getString(11);
		System.out.println("vendorDetails Type vatNo:"+vatNo);
		
		invoiceNo = selectVendorDetailsData.getString(12);
		System.out.println("vendorDetails Type invoiceNo:"+invoiceNo);

	}
	System.out.println("test1...");
	String selectCopNoQuery="select count(*) from dtpo_pocopfile where filetype='COP' and c_order_id='"+orderId+"'";
	selectCopData=statement.executeQuery(selectCopNoQuery);
	while(selectCopData.next()==true){
		
		copCount = selectCopData.getString(1);
		System.out.println("copCount: "+copCount);
	
		int copNo = Integer.parseInt(copCount);
		System.out.println("copNumber: "+copNo);
		
		if(copNo==0){
			 copValue="COP-R001";	
		}
		else{
			String selectCopValueQuery="select max(copno) from dtpo_pocopfile where c_order_id='"+orderId+"'";
		selectCopValueData=statement.executeQuery(selectCopValueQuery);			
		while(selectCopValueData.next()==true){
			previousCop = selectCopValueData.getString(1);
				System.out.println("previouscop: "+previousCop);
		}
				copValue="COP-R00"+copNo;
				System.out.println("copValue: "+copValue);
		}
				
	/*FileInputStream myInput = new FileInputStream(myFileName);
	XSSFWorkbook myWorkBookSheets = new XSSFWorkbook(myInput);
	System.out.println("myWorkBookSheets :"+myWorkBookSheets.getNumberOfSheets());
	//System.out.println("myWorkBook_contract_amt myInput :");
	
	XSSFSheet myFirstSheet = myWorkBookSheets.getSheetAt(1);
	System.out.println("myFirstSheet: "+myFirstSheet);
	//XSSFCell myFirstShetCell = null;
    XSSFSheet mySecondSheet = myWorkBookSheets.getSheetAt(2);
	System.out.println("mySecondSheet: "+mySecondSheet);
	
 	XSSFCell mycell = null;
	
	mycell=mySecondSheet.getRow(7).getCell(3);
	System.out.println("orderdate cell :"+mycell);
	mycell.setCellValue(documentNo+"/"+orderDate);
	
	mycell=mySecondSheet.getRow(5).getCell(5);
	System.out.println("bpname cell :"+mycell);
	mycell.setCellValue(bpartnerName);
	
	mycell=mySecondSheet.getRow(6).getCell(5);
	System.out.println("bpaddr cell :"+mycell);
	mycell.setCellValue(address1+address2+","+city+","+postal+","+countryName);
	
	mycell=mySecondSheet.getRow(7).getCell(5);
	System.out.println("panno cell :"+mycell);
	mycell.setCellValue(panNo);
	
	mycell=mySecondSheet.getRow(8).getCell(5);
	System.out.println("vatno cell :"+mycell);
	mycell.setCellValue(vatNo);
	
	mycell=mySecondSheet.getRow(9).getCell(5);
	System.out.println("invoiceno cell :"+mycell);
	mycell.setCellValue(invoiceNo);
	
	mycell=mySecondSheet.getRow(8).getCell(3);
	System.out.println("grandtotal cell :"+mycell);
	mycell.setCellValue("[$INR] "+grandTotal);
	
	mycell=myFirstSheet.getRow(3).getCell(3);
	System.out.println("copvalue 1stsheet cell :"+mycell);
	mycell.setCellValue(copValue);
	
	mycell=mySecondSheet.getRow(14).getCell(6);
	System.out.println("copvalue 2ndsheet cell :"+mycell);
	mycell.setCellValue(copValue);
	
	mycell=mySecondSheet.getRow(14).getCell(5);
	System.out.println("prevcop cell :"+mycell);
	mycell.setCellValue(previousCop);*/
	
	
	//FileOutputStream fileOut = new FileOutputStream(myFileName);
	//myWorkBookSheets.write(fileOut);
	//fileOut.close();	
	
	
/*}catch (NullPointerException e) {
	e.printStackTrace();

}
		try {
			
			Class.forName(driver);
			conn = DriverManager.getConnection(url, username, password);
			statement = conn.createStatement();*/
			System.out.println("previous cop..");
			String orderScenarioType = "select c_order.em_dtpo_orderscenario from c_order  join dtpo_pocopfile on dtpo_pocopfile.c_order_id=c_ORDER.c_order_id where  dtpo_pocopfile.filename='"+bookname+"'";
			System.out.println("test2...");
			scenarioType = statement.executeQuery(orderScenarioType);
			System.out.println("test3..");
			while(scenarioType.next()==true){
				System.out.println("inside scenario type..");
			orderScenario = scenarioType.getString(1);
			System.out.println("Order Scenario Type :"+orderScenario);
		 	}
			FileInputStream myInput = new FileInputStream(copFileName);
			XSSFWorkbook myWorkBookSheets = new XSSFWorkbook(myInput);
			System.out.println("myWorkBookSheets :"+myWorkBookSheets.getNumberOfSheets());
			//System.out.println("myWorkBook_contract_amt myInput :");
			
			XSSFSheet myFirstSheet = myWorkBookSheets.getSheetAt(1);
			System.out.println("myFirstSheet: "+myFirstSheet);
			
			XSSFCell myFirstShetCell = null;
			System.out.println("myFirstShetCell: "+myFirstShetCell);
            XSSFSheet mySecondSheet = myWorkBookSheets.getSheetAt(2);
			System.out.println("mySecondSheet: "+mySecondSheet);
		 	XSSFCell mycell = null;
			System.out.println("mycell: "+mycell);
			
			
			double firstSheetCopBasic = 0.0;
			double firstSheetCopInsurance = 0.0;
			double firstSheetCopFreight = 0.0;
			double firstSheetCopPkg = 0.0;
			double firstSheetCopOthers = 0.0;
			double firstSheetCopExcise = 0.0;
			double firstSheetCopSubtotal = 0.0;
			double firstSheetCopVAT_CST = 0.0;
			double firstSheetCopLBT = 0.0;
			double firstSheetSupplyTotal = 0.0;
			double firstSheetInstallationService = 0.0;
			double firstSheetInstallationTotal = 0.0;
			double firstSheetCopCess=0.0;
			double firstSheetImportTotal=0.0;
			double copBasic = 0.00;
			double copOthers = 0.00;
			double copInsurance = 0.00;
			double copFreight = 0.00;
			double copPkg = 0.00;
			double copExcise=0.0;
			double copVAT = 0.0;
			double copLBT = 0.0;
			double copSupplyTotal = 0.0;
			double copSupplyVAT = 0.0;
			double copInstallationTotal = 0.0;
			double copInstallationService = 0.0;
			double copCess = 0.0;
			double copImporttotal = 0.0;
			double prevRecColumn1 = 0.00;
			double prevRecColumn2 = 0.00;
			double prevRecColumn3 = 0.00;
			double prevRecColumn4 = 0.00;
			double prevRecColumn5 = 0.00;
			double prevRecColumn6 = 0.00;
			double prevRecColumn7 = 0.00;
			double prevRecColumn8 = 0.00;
			double prevRecColumn9 = 0.00;
			double prevRecColumn10 = 0.00;
			double prevAdvColumn1 = 0.00;
			double prevAdvColumn2 = 0.00;
			double prevAdvColumn3 = 0.00;
			double prevAdvColumn4 = 0.00;
			double prevAdvColumn5 = 0.00;
			double prevAdvColumn6 = 0.00;
			double prevCopBasic = 0.00;
			double prevCopOthers = 0.00;
			double prevCopInsurance = 0.00;
			double prevCopFreight = 0.00;
			double prevCopPkg = 0.00;
			double prevCopExcise = 0.0;
			double prevCopVat = 0.0;
			double prevCopLBT = 0.0;
			double prevSupplyBasic = 0.0;
			double prevSupplyTotal = 0.0;
			double prevImportTotal = 0.0;
			double prevInstallationBasic = 0.0;
			double prevInstallationService = 0.0;
			double prevInstallationOthers = 0.0;
			double prevInstallationTotal = 0.0;
			double prevCopCess = 0.0;
			double cRate =0.0;
			
			//String copNumber = null;
			//myFirstShetCell = myFirstSheet.getRow(3).getCell(3);
			//System.out.println("myFirstShetCell: "+myFirstShetCell);
			//copNumber = myFirstShetCell.toString();
			System.out.println("copValue: "+copValue);
			String copNumber = copValue;
			System.out.println("copNumber: "+copNumber);
			
			//String s=copNumber.substring(copNumber.length()-3,copNumber.length());
			//System.out.println("s: "+s);
			//int prevCopNumber = Integer.parseInt(copNumber) - 1;
			//String prevCop = Integer.toString(prevCopNumber);
			String prevCop = previousCop;
			System.out.println("prevCop :"+prevCop);
			
			selectPrevCOPData = "select dtpo_pocopfile.rec_1stcolumn,dtpo_pocopfile.rec_2ndcolumn," +
					"dtpo_pocopfile.rec_3rdcolumn,dtpo_pocopfile.rec_4thcolumn,dtpo_pocopfile.rec_5thcolumn," +
					"dtpo_pocopfile.rec_6thcolumn,dtpo_pocopfile.rec_7thcolumn,dtpo_pocopfile.rec_8thcolumn,dtpo_pocopfile.rec_9thcolumn," +
					"dtpo_pocopfile.rec_10thcolumn,dtpo_pocopfile.adv_1stcolumn,dtpo_pocopfile.adv_2ndcolumn," +
					"dtpo_pocopfile.adv_3rdcolumn,dtpo_pocopfile.adv_4thcolumn,dtpo_pocopfile.adv_5thcolumn," +
					"dtpo_pocopfile.adv_6thcolumn from dtpo_pocopfile " +
			"left join  dtpo_itemcop on dtpo_itemcop.dtpo_pocopfile_id=dtpo_pocopfile.dtpo_pocopfile_id where copno like '%"+prevCop+"' " +
			"and c_order_id=(select c_order_id from dtpo_pocopfile where filename='"+bookname+"')group by dtpo_pocopfile.rec_1stcolumn," +
					"dtpo_pocopfile.rec_2ndcolumn,dtpo_pocopfile.rec_3rdcolumn,dtpo_pocopfile.rec_4thcolumn," +
					"dtpo_pocopfile.rec_5thcolumn,dtpo_pocopfile.rec_6thcolumn,dtpo_pocopfile.rec_7thcolumn," +
					"dtpo_pocopfile.rec_8thcolumn,dtpo_pocopfile.rec_9thcolumn,dtpo_pocopfile.rec_10thcolumn,dtpo_pocopfile.adv_1stcolumn," +
					"dtpo_pocopfile.adv_2ndcolumn,dtpo_pocopfile.adv_3rdcolumn,dtpo_pocopfile.adv_4thcolumn,dtpo_pocopfile.adv_5thcolumn,dtpo_pocopfile.adv_6thcolumn,dtpo_pocopfile.c_order_id";
			
			orderID = statement.executeQuery(selectPrevCOPData);
			while(orderID.next()==true){
		
				prevRecColumn1 = Double.parseDouble(orderID.getString(1));
				System.out.println("prevRecColumn1 :"+prevRecColumn1);
				prevRecColumn2 = Double.parseDouble(orderID.getString(2));
				System.out.println("prevRecColumn2 :"+prevRecColumn2);
				prevRecColumn3= Double.parseDouble(orderID.getString(3));
				System.out.println("prevRecColumn3 :"+prevRecColumn3);
                prevRecColumn4 = Double.parseDouble(orderID.getString(4));
                System.out.println("prevRecColumn4 :"+prevRecColumn4);
                prevRecColumn5 = Double.parseDouble(orderID.getString(5));
                System.out.println("prevRecColumn5 :"+prevRecColumn5);
				prevRecColumn6 = Double.parseDouble(orderID.getString(6));
				System.out.println("prevRecColumn6 :"+prevRecColumn6);
				prevRecColumn7 = Double.parseDouble(orderID.getString(7));
				System.out.println("prevRecColumn7 :"+prevRecColumn7);
                prevRecColumn8 = Double.parseDouble(orderID.getString(8));
                System.out.println("prevRecColumn8 :"+prevRecColumn8);
                prevRecColumn9 = Double.parseDouble(orderID.getString(9));
                System.out.println("prevRecColumn9 :"+prevRecColumn9);
				prevRecColumn10 = Double.parseDouble(orderID.getString(10));
				System.out.println("prevRecColumn10 :"+prevRecColumn10);
				prevAdvColumn1 = Double.parseDouble(orderID.getString(11));
				System.out.println("prevAdvColumn1 :"+prevAdvColumn1);
                prevAdvColumn2 = Double.parseDouble(orderID.getString(12));
                System.out.println("prevAdvColumn2 :"+prevAdvColumn2);
                prevAdvColumn3 = Double.parseDouble(orderID.getString(13));
                System.out.println("prevAdvColumn3 :"+prevAdvColumn3);
				prevAdvColumn4 = Double.parseDouble(orderID.getString(14));
				System.out.println("prevAdvColumn4 :"+prevAdvColumn4);
				prevAdvColumn5 = Double.parseDouble(orderID.getString(15));
				System.out.println("prevAdvColumn5 :"+prevAdvColumn5);
                prevAdvColumn6 = Double.parseDouble(orderID.getString(16));
                System.out.println("prevAdvColumn6 :"+prevAdvColumn6);
				
				
			
				mycell=mySecondSheet.getRow(23).getCell(5);
				System.out.println("23rd cell :"+mycell);
				mycell.setCellValue(prevRecColumn1);
				
				mycell=mySecondSheet.getRow(24).getCell(5);
				System.out.println("24th cell :"+mycell);
				mycell.setCellValue(prevRecColumn2);
				
				mycell=mySecondSheet.getRow(25).getCell(5);
				System.out.println("25th cell :"+mycell);
				mycell.setCellValue(prevRecColumn3);
				
				mycell=mySecondSheet.getRow(26).getCell(5);
				System.out.println("26th cell :"+mycell);
				mycell.setCellValue(prevRecColumn4);
				
				mycell=mySecondSheet.getRow(27).getCell(5);
				System.out.println("27th cell :"+mycell);
				mycell.setCellValue(prevRecColumn5);
				
				mycell=mySecondSheet.getRow(28).getCell(5);
				System.out.println("28th cell :"+mycell);
				mycell.setCellValue(prevRecColumn6);
				
				mycell=mySecondSheet.getRow(29).getCell(5);
				System.out.println("29th cell :"+mycell);
				mycell.setCellValue(prevRecColumn7);
				
				mycell=mySecondSheet.getRow(30).getCell(5);
				System.out.println("30th cell :"+mycell);
				mycell.setCellValue(prevRecColumn8);
				
				mycell=mySecondSheet.getRow(31).getCell(5);
				System.out.println("31st cell :"+mycell);
				mycell.setCellValue(prevRecColumn9);
				
				mycell=mySecondSheet.getRow(32).getCell(5);
				System.out.println("32nd cell :"+mycell);
				mycell.setCellValue(prevRecColumn10);
				
				mycell=mySecondSheet.getRow(35).getCell(5);
				System.out.println("35th cell :"+mycell);
				mycell.setCellValue(prevAdvColumn1);
				
				mycell=mySecondSheet.getRow(36).getCell(5);
				System.out.println("36th cell :"+mycell);
				mycell.setCellValue(prevAdvColumn2);
				
				mycell=mySecondSheet.getRow(37).getCell(5);
				System.out.println("37th cell :"+mycell);
				mycell.setCellValue(prevAdvColumn3);
				
				mycell=mySecondSheet.getRow(38).getCell(5);
				System.out.println("38th cell :"+mycell);
				mycell.setCellValue(prevAdvColumn4);
				
				mycell=mySecondSheet.getRow(39).getCell(5);
				System.out.println("39th cell :"+mycell);
				mycell.setCellValue(prevAdvColumn5);
				
				mycell=mySecondSheet.getRow(40).getCell(5);
				System.out.println("40th cell :"+mycell);
				mycell.setCellValue(prevAdvColumn6);
				
				
			}

			if(orderScenario.contentEquals("TYPE1")){
				
				XSSFRow cRow = myFirstSheet.getRow(3);
				XSSFCell cCell = cRow.getCell(6);
				cRate = cCell.getNumericCellValue();
				selectCOPBasicData = "select sum(dtpo_itemcop.copbasic),sum(dtpo_itemcop.copinsurance),sum(dtpo_itemcop.copfreight)," +
						"sum(dtpo_itemcop.coppkg_fwdg),sum(dtpo_itemcop.copothers) from dtpo_pocopfile left join  " +
						"dtpo_itemcop on dtpo_itemcop.dtpo_pocopfile_id=dtpo_pocopfile.dtpo_pocopfile_id where " +
						"copno like '%"+prevCop+"' and c_order_id=(select c_order_id from dtpo_pocopfile where filename='"+bookname+"')";
				prevCOPData = statement.executeQuery(selectCOPBasicData);
				while(prevCOPData.next()==true){
					prevCopBasic = Double.parseDouble(prevCOPData.getString(1));
					System.out.println("prevCopBasic :"+prevCopBasic);
					prevCopInsurance = Double.parseDouble(prevCOPData.getString(2));
					System.out.println("prevCopInsurance :"+prevCopInsurance);
					prevCopFreight = Double.parseDouble(prevCOPData.getString(3));
					System.out.println("prevCopFreight :"+prevCopFreight);
					prevCopPkg = Double.parseDouble(prevCOPData.getString(4));
					System.out.println("prevCopPkg :"+prevCopPkg);
					prevCopOthers = Double.parseDouble(prevCOPData.getString(5));
					System.out.println("prevCopOthers :"+prevCopOthers);
					
				}
				mycell=mySecondSheet.getRow(16).getCell(5);
				System.out.println("16th cell :"+mycell);
				mycell.setCellValue(prevCopBasic / cRate);
				
				mycell=mySecondSheet.getRow(17).getCell(5);
				System.out.println("17th cell :"+mycell);
				mycell.setCellValue(prevCopInsurance / cRate);
				
				mycell=mySecondSheet.getRow(18).getCell(5);
				System.out.println("18th cell :"+mycell);
				mycell.setCellValue(prevCopFreight / cRate);
				
				mycell=mySecondSheet.getRow(19).getCell(5);
				System.out.println("19th cell :"+mycell);
				mycell.setCellValue(prevCopPkg / cRate);
				
				mycell=mySecondSheet.getRow(20).getCell(5);
				System.out.println("20th cell :"+mycell);
				mycell.setCellValue(prevCopOthers / cRate);
				
			myFirstShetCell = myFirstSheet.getRow(3).getCell(25);
			firstSheetCopBasic = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopBasic :"+firstSheetCopBasic); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(26);
			firstSheetCopInsurance = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopInsurance :"+firstSheetCopInsurance);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(27);
			firstSheetCopFreight = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopFreight :"+firstSheetCopFreight);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(28);
			firstSheetCopPkg = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopPkg :"+firstSheetCopPkg);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(29);
			firstSheetCopOthers = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopOthers :"+firstSheetCopOthers);
			
			previousTotals = "select sum(copbasic),sum(copothers),sum(copinsurance),sum(copfreight),sum(coppkg_fwdg) from dtpo_itemcop join  dtpo_pocopfile on dtpo_pocopfile.dtpo_pocopfile_id =dtpo_itemcop.dtpo_pocopfile_id where dtpo_pocopfile.c_order_id=(select dtpo_pocopfile.c_order_id from dtpo_pocopfile  where  filename='"+bookname+"' ) and filetype='COP'";

			rs = statement.executeQuery(previousTotals);
			while (rs.next() == true) {
				System.out.println("while condition checking");
			
				    copBasic = firstSheetCopBasic + Double.parseDouble(rs.getString(1));
				    copOthers = firstSheetCopInsurance + Double.parseDouble(rs.getString(2));
				    copInsurance = firstSheetCopFreight + Double.parseDouble(rs.getString(3));
				    copFreight = firstSheetCopPkg + Double.parseDouble(rs.getString(4));
				    copPkg = firstSheetCopOthers + Double.parseDouble(rs.getString(5));
				    
					/*mycell=mySecondSheet.getRow(16).getCell(7);
					System.out.println("current baisc total cell :"+mycell);
					mycell.setCellValue(copBasic);
					
					mycell=mySecondSheet.getRow(17).getCell(7);
					System.out.println("current insurance total cell :"+mycell);
					mycell.setCellValue(copInsurance);
					
					mycell=mySecondSheet.getRow(18).getCell(7);
					System.out.println("current freight total cell :"+mycell);
					mycell.setCellValue(copFreight);
					
					mycell=mySecondSheet.getRow(19).getCell(7);
					System.out.println("current pkag total cell :"+mycell);
					mycell.setCellValue(copPkg);
					
					mycell=mySecondSheet.getRow(20).getCell(7);
					System.out.println("current others total cell :"+mycell);
					mycell.setCellValue(copOthers);*/
					
			}
		}else if(orderScenario.contentEquals("TYPE2")){
			
			selectCOPBasicData = "select sum(dtpo_itemcop.copbasic),sum(dtpo_itemcop.copinsurance),sum(dtpo_itemcop.copfreight)," +
			"sum(dtpo_itemcop.coppkg_fwdg),sum(dtpo_itemcop.copothers),sum(dtpo_itemcop.copexcise),sum(dtpo_itemcop.copvat_cst),sum(dtpo_itemcop.copoctroi_lbt) from dtpo_pocopfile left join  " +
			"dtpo_itemcop on dtpo_itemcop.dtpo_pocopfile_id=dtpo_pocopfile.dtpo_pocopfile_id where " +
			"copno like '%"+prevCop+"' and c_order_id=(select c_order_id from dtpo_pocopfile where filename='"+bookname+"')";
			prevCOPData = statement.executeQuery(selectCOPBasicData);
			while(prevCOPData.next()==true){
					prevCopBasic = Double.parseDouble(prevCOPData.getString(1));
					System.out.println("prevCopBasic :"+prevCopBasic);
					prevCopInsurance = Double.parseDouble(prevCOPData.getString(2));
					System.out.println("prevCopInsurance :"+prevCopInsurance);
					prevCopFreight = Double.parseDouble(prevCOPData.getString(3));
					System.out.println("prevCopFreight :"+prevCopFreight);
					prevCopPkg = Double.parseDouble(prevCOPData.getString(4));
					System.out.println("prevCopPkg :"+prevCopPkg);
					prevCopOthers = Double.parseDouble(prevCOPData.getString(5));
					System.out.println("prevCopOthers :"+prevCopOthers);
					prevCopExcise = Double.parseDouble(prevCOPData.getString(6));
					System.out.println("prevCopExcise :"+prevCopExcise);
					prevCopVat = Double.parseDouble(prevCOPData.getString(7));
					System.out.println("prevCopVat :"+prevCopVat);
					prevCopLBT = Double.parseDouble(prevCOPData.getString(8));
					System.out.println("prevCopLBT :"+prevCopLBT);
		
			}  
 	
			mycell=mySecondSheet.getRow(16).getCell(5);
			System.out.println("16th cell :"+mycell);
			mycell.setCellValue(prevCopBasic+prevCopExcise);
			
			mycell=mySecondSheet.getRow(17).getCell(5);
			System.out.println("17th cell :"+mycell);
			mycell.setCellValue(prevCopVat);
			
			mycell=mySecondSheet.getRow(18).getCell(5);
			System.out.println("18th cell :"+mycell);
			mycell.setCellValue(prevCopFreight);
			
			mycell=mySecondSheet.getRow(19).getCell(5);
			System.out.println("19th cell :"+mycell);
			mycell.setCellValue(prevCopPkg + prevCopInsurance + prevCopLBT);
			
			mycell=mySecondSheet.getRow(20).getCell(5);
			System.out.println("20th cell :"+mycell);
			mycell.setCellValue(prevCopOthers);
			
			
			myFirstShetCell = myFirstSheet.getRow(3).getCell(29);
			firstSheetCopBasic = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopBasic :"+firstSheetCopBasic);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(30);
			firstSheetCopExcise = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopExcise :"+firstSheetCopExcise); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(31);
			firstSheetCopSubtotal = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopSubtotal :"+firstSheetCopSubtotal); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(32);
			firstSheetCopVAT_CST = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopVAT_CST :"+firstSheetCopVAT_CST); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(33);
			firstSheetCopLBT = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopLBT :"+firstSheetCopLBT); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(34);
			firstSheetCopInsurance = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopInsurance :"+firstSheetCopInsurance);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(35);
			firstSheetCopFreight = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopFreight :"+firstSheetCopFreight);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(36);
			firstSheetCopPkg = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopPkg :"+firstSheetCopPkg);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(37);
			firstSheetCopOthers = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopOthers :"+firstSheetCopOthers);
			
			previousTotals = "select sum(copbasic),sum(copothers),sum(copinsurance),sum(copfreight),sum(coppkg_fwdg),sum(copexcise),sum(copvat_cst),sum(copoctroi_lbt) from dtpo_itemcop join  dtpo_pocopfile on dtpo_pocopfile.dtpo_pocopfile_id =dtpo_itemcop.dtpo_pocopfile_id where dtpo_pocopfile.c_order_id=(select dtpo_pocopfile.c_order_id from dtpo_pocopfile  where  filename='"+bookname+"' ) and filetype='COP'";
			rs = statement.executeQuery(previousTotals);
			while (rs.next() == true) {

			    copBasic = firstSheetCopBasic + Double.parseDouble(rs.getString(1));
			    copOthers = firstSheetCopInsurance + Double.parseDouble(rs.getString(2));
			    copInsurance = firstSheetCopFreight + Double.parseDouble(rs.getString(3));
			    copFreight = firstSheetCopPkg + Double.parseDouble(rs.getString(4));
			    copPkg = firstSheetCopOthers + Double.parseDouble(rs.getString(5));
			    copExcise = firstSheetCopExcise + Double.parseDouble(rs.getString(6));
				copVAT = firstSheetCopVAT_CST + Double.parseDouble(rs.getString(7));
				copLBT = firstSheetCopLBT + Double.parseDouble(rs.getString(8));
				
				/*mycell=mySecondSheet.getRow(16).getCell(7);
				
				mycell.setCellValue(copBasic);
				
				mycell=mySecondSheet.getRow(17).getCell(7);
				
				mycell.setCellValue(copExcise);
				
				mycell=mySecondSheet.getRow(18).getCell(7);
				
				mycell.setCellValue(copVAT);
				
				mycell=mySecondSheet.getRow(19).getCell(7);
				
				mycell.setCellValue(copPkg + copInsurance + copLBT + copFreight);
				
				mycell=mySecondSheet.getRow(20).getCell(7);
				
				mycell.setCellValue(copOthers);*/
			}
			
		}else if(orderScenario.contentEquals("TYPE3")){
			selectCOPBasicData = "select sum(dtpo_itemcop.copsupplybasic),sum(dtpo_itemcop.copexcise)," +
					"sum(dtpo_itemcop.copvat_cst),sum(dtpo_itemcop.copothers),sum(dtpo_itemcop.copsupplytotal)," +
					"sum(dtpo_itemcop.copinstallationbasic),sum(dtpo_itemcop.copinstallationservice)," +
					"sum(dtpo_itemcop.copinstallationothers),sum(dtpo_itemcop.copinstallationtotal) " +
					"from dtpo_pocopfile left join  dtpo_itemcop on dtpo_itemcop.dtpo_pocopfile_id=dtpo_pocopfile.dtpo_pocopfile_id where " +
					"copno like '%"+prevCop+"' and c_order_id=(select c_order_id from dtpo_pocopfile where filename='"+bookname+"')";
			prevCOPData = statement.executeQuery(selectCOPBasicData);
			while(prevCOPData.next()==true){
					prevCopBasic = Double.parseDouble(prevCOPData.getString(1));
					System.out.println("prevCopSupplyBasic :"+prevCopBasic);
					prevCopExcise = Double.parseDouble(prevCOPData.getString(2));
					System.out.println("prevCopExcise :"+prevCopExcise);
					prevCopVat = Double.parseDouble(prevCOPData.getString(3));
					System.out.println("prevCopVat :"+prevCopVat);
					prevCopOthers = Double.parseDouble(prevCOPData.getString(4));
					System.out.println("prevCopOthers :"+prevCopOthers);
					prevSupplyTotal = Double.parseDouble(prevCOPData.getString(5));
					System.out.println("prevSupplyTotal :"+prevSupplyTotal);
					prevInstallationBasic = Double.parseDouble(prevCOPData.getString(6));
					System.out.println("prevInstallationBasic :"+prevInstallationBasic);
					prevInstallationService = Double.parseDouble(prevCOPData.getString(7));
					System.out.println("prevInstallationService :"+prevInstallationService);
					prevInstallationOthers = Double.parseDouble(prevCOPData.getString(8));
					System.out.println("prevInstallationOthers :"+prevInstallationOthers);
					prevInstallationTotal = Double.parseDouble(prevCOPData.getString(8));
					System.out.println("prevInstallationTotal :"+prevInstallationTotal);
		
			}  
 	
			mycell=mySecondSheet.getRow(16).getCell(5);
			System.out.println("16th cell :"+mycell);
		//	mycell.setCellValue(prevSupplyTotal - prevCopVat);
			mycell.setCellValue(prevCopBasic + prevCopExcise);
			
			mycell=mySecondSheet.getRow(17).getCell(5);
			System.out.println("17th cell :"+mycell);
			//mycell.setCellValue(prevInstallationTotal - prevInstallationService);
			mycell.setCellValue(prevInstallationBasic);
			
			mycell=mySecondSheet.getRow(18).getCell(5);
			System.out.println("18th cell :"+mycell);
			mycell.setCellValue(prevCopVat);
			
			mycell=mySecondSheet.getRow(19).getCell(5);
			System.out.println("19th cell :"+mycell);
			mycell.setCellValue(prevInstallationService);
			
			mycell=mySecondSheet.getRow(20).getCell(5);
			System.out.println("20th cell :"+mycell);
			mycell.setCellValue(prevInstallationOthers+prevCopOthers);
			
			
			myFirstShetCell = myFirstSheet.getRow(3).getCell(35);
			firstSheetCopVAT_CST = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopVAT_CST :"+firstSheetCopVAT_CST);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(37);
			firstSheetSupplyTotal = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetSupplyTotal :"+firstSheetSupplyTotal); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(39);
			firstSheetInstallationService = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetInstallationService :"+firstSheetInstallationService); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(41);
			firstSheetInstallationTotal = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetInstallationTotal :"+firstSheetInstallationTotal); 
			
			/*myFirstShetCell = myFirstSheet.getRow(3).getCell(33);
			firstSheetCopLBT = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopLBT :"+firstSheetCopLBT); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(34);
			firstSheetCopInsurance = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopInsurance :"+firstSheetCopInsurance);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(35);
			firstSheetCopFreight = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopFreight :"+firstSheetCopFreight);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(36);
			firstSheetCopPkg = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopPkg :"+firstSheetCopPkg);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(37);
			firstSheetCopOthers = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopOthers :"+firstSheetCopOthers);*/
			
			previousTotals = "select sum(copsupplytotal),sum(copvat_cst),sum(copinstallationtotal),sum(copinstallationservice) from dtpo_itemcop join  dtpo_pocopfile on dtpo_pocopfile.dtpo_pocopfile_id =dtpo_itemcop.dtpo_pocopfile_id where dtpo_pocopfile.c_order_id=(select dtpo_pocopfile.c_order_id from dtpo_pocopfile  where  filename='"+bookname+"' ) and filetype='COP'";
			rs = statement.executeQuery(previousTotals);
			while (rs.next() == true) {

			    copSupplyTotal = firstSheetSupplyTotal + Double.parseDouble(rs.getString(1));
			    copSupplyVAT = firstSheetCopInsurance + Double.parseDouble(rs.getString(2));
			    copInstallationTotal = firstSheetCopFreight + Double.parseDouble(rs.getString(3));
			    copInstallationService = firstSheetCopPkg + Double.parseDouble(rs.getString(4));
			    
				
				/*mycell=mySecondSheet.getRow(16).getCell(7);
				
				mycell.setCellValue(copSupplyTotal - copSupplyVAT);
				
				mycell=mySecondSheet.getRow(17).getCell(7);
				
				mycell.setCellValue(copInstallationTotal - copInstallationService);
				
				mycell=mySecondSheet.getRow(18).getCell(7);
				
				mycell.setCellValue(copVAT);
				
				mycell=mySecondSheet.getRow(19).getCell(7);
				
				mycell.setCellValue(copSupplyVAT);
				
				mycell=mySecondSheet.getRow(20).getCell(7);
				
				mycell.setCellValue(copInstallationService);*/
			}
			
		}else if(orderScenario.contentEquals("TYPE4")){
			selectCOPBasicData = "select sum(dtpo_itemcop.copbasic),sum(dtpo_itemcop.copvat_cst),sum(dtpo_itemcop.copothers)," +
					"sum(dtpo_itemcop.copinstallationservice),sum(dtpo_itemcop.copcess) from dtpo_pocopfile left join  " +
					"dtpo_itemcop on dtpo_itemcop.dtpo_pocopfile_id=dtpo_pocopfile.dtpo_pocopfile_id where " +
					"copno like '%"+prevCop+"' and c_order_id=(select c_order_id from dtpo_pocopfile where " +
							"filename='"+bookname+"')";

				prevCOPData = statement.executeQuery(selectCOPBasicData);
				while(prevCOPData.next()==true){
					prevCopBasic = Double.parseDouble(prevCOPData.getString(1));
					System.out.println("prevCopSupplyBasic :"+prevCopBasic);
					
					prevCopVat = Double.parseDouble(prevCOPData.getString(2));
					System.out.println("prevCopVat :"+prevCopVat);
					
					prevCopOthers = Double.parseDouble(prevCOPData.getString(3));
					System.out.println("prevCopOthers :"+prevCopOthers);
			
			
					prevInstallationService = Double.parseDouble(prevCOPData.getString(4));
					System.out.println("prevInstallationService :"+prevInstallationService);
			
					prevCopCess = Double.parseDouble(prevCOPData.getString(5));
					System.out.println("prevCopCess :"+prevCopCess);
			

				}  

					mycell=mySecondSheet.getRow(16).getCell(5);
					System.out.println("16th cell :"+mycell);
					mycell.setCellValue(prevCopBasic);
	
					mycell=mySecondSheet.getRow(17).getCell(5);
					System.out.println("17th cell :"+mycell);
					mycell.setCellValue(prevCopVat);
	
					mycell=mySecondSheet.getRow(18).getCell(5);
					System.out.println("18th cell :"+mycell);
					mycell.setCellValue(prevInstallationService);
	
					mycell=mySecondSheet.getRow(19).getCell(5);
					System.out.println("19th cell :"+mycell);
					mycell.setCellValue(prevCopCess);
	
					mycell=mySecondSheet.getRow(20).getCell(5);
					System.out.println("20th cell :"+mycell);
					mycell.setCellValue(prevCopOthers);
	
	
					/*myFirstShetCell = myFirstSheet.getRow(3).getCell(25);
					firstSheetCopBasic = Double.parseDouble(myFirstShetCell.getRawValue());
					System.out.println("firstSheetCopBasic :"+firstSheetCopBasic);
					myFirstShetCell = myFirstSheet.getRow(3).getCell(26);
					firstSheetCopVAT_CST = Double.parseDouble(myFirstShetCell.getRawValue());
					System.out.println("firstSheetCopVAT_CST :"+firstSheetCopVAT_CST); 
					myFirstShetCell = myFirstSheet.getRow(3).getCell(27);
					firstSheetInstallationService = Double.parseDouble(myFirstShetCell.getRawValue());
					System.out.println("firstSheetInstallationService :"+firstSheetInstallationService); 
					myFirstShetCell = myFirstSheet.getRow(3).getCell(28);
					firstSheetCopCess = Double.parseDouble(myFirstShetCell.getRawValue());
					System.out.println("firstSheetCopCess :"+firstSheetCopCess); 
					myFirstShetCell = myFirstSheet.getRow(3).getCell(29);
					firstSheetCopOthers = Double.parseDouble(myFirstShetCell.getRawValue());
					System.out.println("firstSheetCopOthers :"+firstSheetCopOthers); */
	
	
					previousTotals = "select sum(copbasic),sum(copvat_cst),sum(copinstallationservice),sum(copcess),sum(copothers)  from dtpo_itemcop join  dtpo_pocopfile on dtpo_pocopfile.dtpo_pocopfile_id =dtpo_itemcop.dtpo_pocopfile_id where dtpo_pocopfile.c_order_id=(select dtpo_pocopfile.c_order_id from dtpo_pocopfile  where  filename='"+bookname+"' ) and filetype='COP'";
					rs = statement.executeQuery(previousTotals);
						while (rs.next() == true) {

							copBasic = firstSheetCopBasic + Double.parseDouble(rs.getString(1));
							copSupplyVAT = firstSheetCopVAT_CST + Double.parseDouble(rs.getString(2));
							copInstallationService = firstSheetInstallationService + Double.parseDouble(rs.getString(3));
							copCess = firstSheetCopCess + Double.parseDouble(rs.getString(4));
							copOthers = firstSheetCopOthers + Double.parseDouble(rs.getString(5));
		
								/*mycell=mySecondSheet.getRow(16).getCell(7);
		
								mycell.setCellValue(copBasic);
		
								mycell=mySecondSheet.getRow(17).getCell(7);
		
								mycell.setCellValue(copSupplyVAT);
								
								mycell=mySecondSheet.getRow(18).getCell(7);
		
								mycell.setCellValue(copInstallationService);
		
								mycell=mySecondSheet.getRow(19).getCell(7);
		
								mycell.setCellValue(copCess);
		
								mycell=mySecondSheet.getRow(20).getCell(7);
		
								mycell.setCellValue(copOthers);*/
						}
	
		}else if(orderScenario.contentEquals("TYPE5")){
			selectCOPBasicData = "select sum(dtpo_itemcop.copbasic),sum(dtpo_itemcop.copvat_cst),sum(dtpo_itemcop.copothers) from dtpo_pocopfile left join  " +
			"dtpo_itemcop on dtpo_itemcop.dtpo_pocopfile_id=dtpo_pocopfile.dtpo_pocopfile_id where " +
			"copno like '%"+prevCop+"' and c_order_id=(select c_order_id from dtpo_pocopfile where " +
					"filename='"+bookname+"')";

		prevCOPData = statement.executeQuery(selectCOPBasicData);
		while(prevCOPData.next()==true){
			prevCopBasic = Double.parseDouble(prevCOPData.getString(1));
			System.out.println("prevCopSupplyBasic :"+prevCopBasic);
			
			prevCopVat = Double.parseDouble(prevCOPData.getString(2));
			System.out.println("prevCopVat :"+prevCopVat);
			
			prevCopOthers = Double.parseDouble(prevCOPData.getString(3));
			System.out.println("prevCopOthers :"+prevCopOthers);
	

		}  

			mycell=mySecondSheet.getRow(16).getCell(5);
			System.out.println("16th cell :"+mycell);
			mycell.setCellValue(prevCopBasic);

			mycell=mySecondSheet.getRow(17).getCell(5);
			System.out.println("17th cell :"+mycell);
			mycell.setCellValue(prevCopVat);

			mycell=mySecondSheet.getRow(18).getCell(5);
			System.out.println("18th cell :"+mycell);
			mycell.setCellValue(prevCopOthers);

			
			myFirstShetCell = myFirstSheet.getRow(3).getCell(23);
			firstSheetCopBasic = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopBasic :"+firstSheetCopBasic);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(24);
			firstSheetCopVAT_CST = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopVAT_CST :"+firstSheetCopVAT_CST); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(25);
			firstSheetCopOthers = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopOthers :"+firstSheetCopOthers); 
			
			previousTotals = "select sum(copbasic),sum(copvat_cst),sum(copothers)  from dtpo_itemcop join  dtpo_pocopfile on dtpo_pocopfile.dtpo_pocopfile_id =dtpo_itemcop.dtpo_pocopfile_id where dtpo_pocopfile.c_order_id=(select dtpo_pocopfile.c_order_id from dtpo_pocopfile  where  filename='"+bookname+"' ) and filetype='COP'";
			rs = statement.executeQuery(previousTotals);
				while (rs.next() == true) {

					copBasic = firstSheetCopBasic + Double.parseDouble(rs.getString(1));
					copSupplyVAT = firstSheetCopVAT_CST + Double.parseDouble(rs.getString(2));
					copOthers = firstSheetCopOthers + Double.parseDouble(rs.getString(3));
					

						/*mycell=mySecondSheet.getRow(16).getCell(7);

						mycell.setCellValue(copBasic);

						mycell=mySecondSheet.getRow(17).getCell(7);

						mycell.setCellValue(copSupplyVAT);
						
						mycell=mySecondSheet.getRow(18).getCell(7);

						mycell.setCellValue(copOthers);*/
		
				}
	
			
		}else if(orderScenario.contentEquals("TYPE6")){
			
			selectCOPBasicData = "select sum(dtpo_itemcop.copimportedtotal),sum(dtpo_itemcop.copsupplytotal)," +
					"sum(dtpo_itemcop.copinstallationtotal),sum(dtpo_itemcop.copvat_cst),sum(dtpo_itemcop.copinstallationservice) from dtpo_pocopfile left join  " +
			"dtpo_itemcop on dtpo_itemcop.dtpo_pocopfile_id=dtpo_pocopfile.dtpo_pocopfile_id where " +
			"copno like '%"+prevCop+"' and c_order_id=(select c_order_id from dtpo_pocopfile where " +
					"filename='"+bookname+"')";

		prevCOPData = statement.executeQuery(selectCOPBasicData);
		while(prevCOPData.next()==true){
			prevImportTotal = Double.parseDouble(prevCOPData.getString(1));
			System.out.println("prevImportTotal :"+prevImportTotal);
			prevSupplyTotal = Double.parseDouble(prevCOPData.getString(2));
			System.out.println("prevSupplyTotal :"+prevSupplyTotal);
			prevInstallationTotal = Double.parseDouble(prevCOPData.getString(3));
			System.out.println("prevInstallationTotal :"+prevInstallationTotal);
			prevCopVat = Double.parseDouble(prevCOPData.getString(4));
			System.out.println("prevCopVat :"+prevCopVat);
			prevInstallationService = Double.parseDouble(prevCOPData.getString(5));
			System.out.println("prevInstallationService :"+prevInstallationService);
			}  

			mycell=mySecondSheet.getRow(16).getCell(5);
			System.out.println("16th cell :"+mycell);
			mycell.setCellValue(prevImportTotal);

			mycell=mySecondSheet.getRow(17).getCell(5);
			System.out.println("17th cell :"+mycell);
			mycell.setCellValue(prevSupplyTotal-prevCopVat);

			mycell=mySecondSheet.getRow(18).getCell(5);
			System.out.println("18th cell :"+mycell);
			mycell.setCellValue(prevCopVat);
			
			mycell=mySecondSheet.getRow(19).getCell(5);
			System.out.println("18th cell :"+mycell);
			mycell.setCellValue(prevInstallationTotal-prevInstallationService);
			
			mycell=mySecondSheet.getRow(20).getCell(5);
			System.out.println("18th cell :"+mycell);
			mycell.setCellValue(prevInstallationService);

			
			myFirstShetCell = myFirstSheet.getRow(3).getCell(44);
			firstSheetImportTotal = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetImportTotal :"+firstSheetImportTotal);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(48);
			firstSheetCopVAT_CST = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetCopVAT_CST :"+firstSheetCopVAT_CST); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(50);
			firstSheetSupplyTotal = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetSupplyTotal :"+firstSheetSupplyTotal); 
			myFirstShetCell = myFirstSheet.getRow(3).getCell(52);
			firstSheetInstallationService = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetInstallationService :"+firstSheetInstallationService);
			myFirstShetCell = myFirstSheet.getRow(3).getCell(54);
			firstSheetInstallationTotal = Double.parseDouble(myFirstShetCell.getRawValue());
			System.out.println("firstSheetInstallationTotal :"+firstSheetInstallationTotal);
			
			previousTotals = "select sum(copimportedtotal),sum(copvat_cst),sum(copsupplytotal),sum(copinstallationservice),sum(copinstallationtotal)  from dtpo_itemcop join  dtpo_pocopfile on dtpo_pocopfile.dtpo_pocopfile_id =dtpo_itemcop.dtpo_pocopfile_id where dtpo_pocopfile.c_order_id=(select dtpo_pocopfile.c_order_id from dtpo_pocopfile  where  filename='"+bookname+"' ) and filetype='COP'";
			rs = statement.executeQuery(previousTotals);
				while (rs.next() == true) {

					copImporttotal = firstSheetImportTotal + Double.parseDouble(rs.getString(1));
					copSupplyVAT = firstSheetCopVAT_CST + Double.parseDouble(rs.getString(2));
					copSupplyTotal = firstSheetSupplyTotal + Double.parseDouble(rs.getString(3));
					copInstallationService = firstSheetInstallationService + Double.parseDouble(rs.getString(4));
					copInstallationTotal = firstSheetInstallationTotal + Double.parseDouble(rs.getString(5));
				
						
				}
	
		}
			//FileInputStream myInput = new FileInputStream(myFileName);
			//XSSFWorkbook myWorkBookSheets = new XSSFWorkbook(myInput);
			//System.out.println("myWorkBookSheets :"+myWorkBookSheets.getNumberOfSheets());
			//System.out.println("myWorkBook_contract_amt myInput :");
			
			//XSSFSheet myFirstSheet = myWorkBookSheets.getSheetAt(1);
			//System.out.println("myFirstSheet: "+myFirstSheet);
			//XSSFCell myFirstShetCell = null;
		    //XSSFSheet mySecondSheet = myWorkBookSheets.getSheetAt(2);
			//System.out.println("mySecondSheet: "+mySecondSheet);
			
		 	//XSSFCell mycell = null;
			
			mycell=mySecondSheet.getRow(7).getCell(3);
			System.out.println("orderdate cell :"+mycell);
			mycell.setCellValue(documentNo+"/"+orderDate);
			
			mycell=mySecondSheet.getRow(5).getCell(5);
			System.out.println("bpname cell :"+mycell);
			mycell.setCellValue(bpartnerName);
			
			mycell=mySecondSheet.getRow(6).getCell(5);
			System.out.println("bpaddr cell :"+mycell);
			mycell.setCellValue(address1+address2+","+city+","+postal+","+countryName);
			
			mycell=mySecondSheet.getRow(7).getCell(5);
			System.out.println("panno cell :"+mycell);
			mycell.setCellValue(panNo);
			
			mycell=mySecondSheet.getRow(8).getCell(5);
			System.out.println("vatno cell :"+mycell);
			mycell.setCellValue(vatNo);
			
			mycell=mySecondSheet.getRow(9).getCell(5);
			System.out.println("invoiceno cell :"+mycell);
			mycell.setCellValue(invoiceNo);
			
			mycell=mySecondSheet.getRow(8).getCell(3);
			System.out.println("grandtotal cell :"+mycell);
			mycell.setCellValue("[$INR] "+grandTotal);
			
			mycell=myFirstSheet.getRow(3).getCell(3);
			System.out.println("copvalue 1stsheet cell :"+mycell);
			mycell.setCellValue(copValue);
			
			mycell=mySecondSheet.getRow(14).getCell(6);
			System.out.println("copvalue 2ndsheet cell :"+mycell);
			mycell.setCellValue(copValue);
			
			mycell=mySecondSheet.getRow(14).getCell(5);
			System.out.println("prevcop cell :"+mycell);
			mycell.setCellValue(previousCop);
		
			FileOutputStream fileOut = new FileOutputStream(copFileName);
			myWorkBookSheets.write(fileOut);
			fileOut.close();
			Messagebox.show("Previous COP Data Imported Successfully Re-Open File On Import SpreadSheet Window "+bookname);
			
			//Executions.sendRedirect("/copfile.zul?filename="+bookname);
			
	}
		}
		catch (NullPointerException e) {
			e.printStackTrace();

		}
	}
}
