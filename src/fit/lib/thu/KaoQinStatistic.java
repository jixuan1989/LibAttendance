/**
 * 
 */
package fit.lib.thu;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFeatures;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * @version 1.0
 * @author zhuoan
 * @version 2.0
 * @author huangxiangdong 傻x了， 下一版本的实现者，应该参考这样的思路：先一口气把数据清理好，然后再判断。。现在的实现太累了。。。
 */
public class KaoQinStatistic {

	/**
	 * 生成Excel文件
	 * 
	 * @param path
	 *            文件路径
	 * @param sheetName
	 *            工作表名称
	 * @param dataTitles
	 *            数据标题
	 */
	public void createExcelFile(String path, String sheetName,
			String[] dataTitles) {

		WritableWorkbook workbook;
		try {
			OutputStream os = new FileOutputStream(path);
			workbook = Workbook.createWorkbook(os);

			WritableSheet sheet = workbook.createSheet(sheetName, 0); // 添加第一个工作表
			initialSheetSetting(sheet);

			Label label;
			for (int i = 0; i < dataTitles.length; i++) {
				// Label(列号,行号,内容,风格)
				label = new Label(i, 0, dataTitles[i], getTitleCellFormat());
				sheet.addCell(label);
			}

			// 插入一行
			insertRowData(sheet, 1, new String[] { "200201001", "张三", "100",
					"60", "100", "260" },
					getDataCellFormat(CellType.STRING_FORMULA));

			// 一个一个插入行
			label = new Label(0, 2, "200201002",
					getDataCellFormat(CellType.STRING_FORMULA));
			sheet.addCell(label);

			label = new Label(1, 2, "李四",
					getDataCellFormat(CellType.STRING_FORMULA));
			sheet.addCell(label);

			insertOneCellData(sheet, 2, 2, 70.5,
					getDataCellFormat(CellType.NUMBER));
			insertOneCellData(sheet, 3, 2, 90.523,
					getDataCellFormat(CellType.NUMBER));
			insertOneCellData(sheet, 4, 2, 60.5,
					getDataCellFormat(CellType.NUMBER));

			insertFormula(sheet, 5, 2, "C3+D3+E3",
					getDataCellFormat(CellType.NUMBER_FORMULA));

			// 插入日期
			mergeCellsAndInsertData(sheet, 0, 3, 5, 3, new Date(),
					getDataCellFormat(CellType.DATE));

			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 初始化表格属性
	 * 
	 * @param sheet
	 */
	public void initialSheetSetting(WritableSheet sheet) {
		try {
			// sheet.getSettings().setProtected(true); //设置xls的保护，单元格为只读的
			sheet.getSettings().setDefaultColumnWidth(10); // 设置列的默认宽度
			// sheet.setRowView(2,false);//行高自动扩展
			// setRowView(int row, int height);–行高
			// setColumnView(int col,int width); –列宽
			sheet.setColumnView(0, 10);// 设置第一列宽度
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 插入公式
	 * 
	 * @param sheet
	 * @param col
	 * @param row
	 * @param formula
	 * @param format
	 */
	public void insertFormula(WritableSheet sheet, Integer col, Integer row,
			String formula, WritableCellFormat format) {
		try {
			Formula f = new Formula(col, row, formula, format);
			sheet.addCell(f);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 插入一行数据
	 * 
	 * @param sheet
	 *            工作表
	 * @param row
	 *            行号
	 * @param content
	 *            内容
	 * @param format
	 *            风格
	 */
	public static  void insertRowData(WritableSheet sheet, Integer row,
			String[] dataArr, WritableCellFormat format) {
		try {
			Label label;
			for (int i = 0; i < dataArr.length; i++) {
				label = new Label(i, row, dataArr[i], format);
				sheet.addCell(label);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 插入单元格数据
	 * 
	 * @param sheet
	 * @param col
	 * @param row
	 * @param data
	 */
	public static void insertOneCellData(WritableSheet sheet, Integer col,
			Integer row, Object data, WritableCellFormat format) {
		try {
			if (data instanceof Double) {
				jxl.write.Number labelNF = new jxl.write.Number(col, row,
						(Double) data, format);
				sheet.addCell(labelNF);
			} else if (data instanceof Boolean) {
				jxl.write.Boolean labelB = new jxl.write.Boolean(col, row,
						(Boolean) data, format);
				sheet.addCell(labelB);
			} else if (data instanceof Date) {
				jxl.write.DateTime labelDT = new jxl.write.DateTime(col, row,
						(Date) data, format);
				sheet.addCell(labelDT);
				setCellComments(labelDT, "这是个创建表的日期说明！");
			} else {
				Label label = new Label(col, row, data.toString(), format);
				sheet.addCell(label);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * 合并单元格，并插入数据
	 * 
	 * @param sheet
	 * @param col_start
	 * @param row_start
	 * @param col_end
	 * @param row_end
	 * @param data
	 * @param format
	 */
	public void mergeCellsAndInsertData(WritableSheet sheet, Integer col_start,
			Integer row_start, Integer col_end, Integer row_end, Object data,
			WritableCellFormat format) {
		try {
			sheet.mergeCells(col_start, row_start, col_end, row_end);// 左上角到右下角
			insertOneCellData(sheet, col_start, row_start, data, format);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * 给单元格加注释
	 * 
	 * @param label
	 * @param comments
	 */
	public  static void  setCellComments(Object label, String comments) {
		WritableCellFeatures cellFeatures = new WritableCellFeatures();
		cellFeatures.setComment(comments);
		if (label instanceof jxl.write.Number) {
			jxl.write.Number num = (jxl.write.Number) label;
			num.setCellFeatures(cellFeatures);
		} else if (label instanceof jxl.write.Boolean) {
			jxl.write.Boolean bool = (jxl.write.Boolean) label;
			bool.setCellFeatures(cellFeatures);
		} else if (label instanceof jxl.write.DateTime) {
			jxl.write.DateTime dt = (jxl.write.DateTime) label;
			dt.setCellFeatures(cellFeatures);
		} else {
			Label _label = (Label) label;
			_label.setCellFeatures(cellFeatures);
		}
	}

	/**
	 * 读取excel
	 * 
	 * @param inputFile
	 * @param inputFileSheetIndex
	 * @throws Exception
	 */
	public ArrayList<String> readDataFromExcel(File inputFile,
			int inputFileSheetIndex) {

		ArrayList<String> list = new ArrayList<String>();
		Workbook book = null;
		Cell cell = null;
		WorkbookSettings setting = new WorkbookSettings();
		java.util.Locale locale = new java.util.Locale("zh", "CN");
		setting.setLocale(locale);
		setting.setEncoding("ISO-8859-1");
		try {
			book = Workbook.getWorkbook(inputFile, setting);
		} catch (Exception e) {
			e.printStackTrace();
		}

		Sheet sheet = book.getSheet(inputFileSheetIndex);
		for (int rowIndex = 0; rowIndex < sheet.getRows(); rowIndex++) {// 行
			for (int colIndex = 0; colIndex < sheet.getColumns(); colIndex++) {// 列
				cell = sheet.getCell(colIndex, rowIndex);
				// System.out.println(cell.getContents());
				list.add(cell.getContents());
			}
		}
		book.close();

		return list;
	}

	/**
	 * 得到数据表头格式
	 * 
	 * @return
	 */
	public WritableCellFormat getTitleCellFormat() {
		WritableCellFormat wcf = null;
		try {
			// 字体样式
			WritableFont wf = new WritableFont(WritableFont.TIMES, 12,
					WritableFont.NO_BOLD, false);// 最后一个为是否italic
			wf.setColour(Colour.RED);
			wcf = new WritableCellFormat(wf);
			// 对齐方式
			wcf.setAlignment(Alignment.CENTRE);
			wcf.setVerticalAlignment(VerticalAlignment.CENTRE);
			// 边框
			wcf.setBorder(Border.ALL, BorderLineStyle.THIN);

			// 背景色
			wcf.setBackground(Colour.GREY_25_PERCENT);
		} catch (WriteException e) {
			e.printStackTrace();
		}
		return wcf;
	}

	/**
	 * 得到数据格式
	 * 
	 * @return
	 */
	public static WritableCellFormat getDataCellFormat(CellType type) {
		WritableCellFormat wcf = null;
		try {
			// 字体样式
			if (type == CellType.NUMBER || type == CellType.NUMBER_FORMULA) {// 数字
				NumberFormat nf = new NumberFormat("#.00");
				wcf = new WritableCellFormat(nf);
				// 背景色
				wcf.setBackground(Colour.YELLOW);
			} else if (type == CellType.DATE || type == CellType.DATE_FORMULA) {// 日期
				jxl.write.DateFormat df = new jxl.write.DateFormat(
						"yyyy-MM-dd hh:mm:ss");
				wcf = new jxl.write.WritableCellFormat(df);
				// 背景色
				// wcf.setBackground(Colour.YELLOW);
			} else {
				WritableFont wf = new WritableFont(WritableFont.TIMES, 10,
						WritableFont.NO_BOLD, false);// 最后一个为是否italic
				wcf = new WritableCellFormat(wf);
				// 背景色
				if (type == CellType.LABEL)
					wcf.setBackground(Colour.RED);
				else if(type==CellType.STRING_FORMULA){
					wcf.setBackground(Colour.WHITE);
				}
			}
			// 对齐方式
			wcf.setAlignment(Alignment.CENTRE);
			wcf.setVerticalAlignment(VerticalAlignment.CENTRE);
			// 边框
			wcf.setBorder(Border.LEFT, BorderLineStyle.THIN);
			wcf.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
			wcf.setBorder(Border.RIGHT, BorderLineStyle.THIN);

			wcf.setWrap(true);// 自动换行

		} catch (WriteException e) {
			e.printStackTrace();
		}
		return wcf;
	}

	/**
	 * 打开文件看看
	 * 
	 * @param exePath
	 * @param filePath
	 */
	public void openExcel(String exePath, String filePath) {
		Runtime r = Runtime.getRuntime();
		String cmd[] = { exePath, filePath };
		try {
			r.exec(cmd);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void test() {

		String[] titles = { "学号", "姓名", "语文", "数学", "英语", "总分" };
		KaoQinStatistic jxl = new KaoQinStatistic();
		String filePath = "G:/InOutData.xls";
		jxl.createExcelFile(filePath, "成绩单", titles);
		System.out.println(jxl.readDataFromExcel(new File(filePath), 0));
		jxl.openExcel("c:/Program Files/Microsoft Office/OFFICE12/EXCEL.EXE",
				filePath);
	}



	private static int parserColumnName2Int(String columnName) {
		char[] array = columnName.toCharArray();
		int result = 0;
		for (int i = array.length - 1; i >= 0; i--) {
			result += Math.pow(26, array.length - 1 - i) * (array[i] - 'A' + 1);
		}
		return result - 1;
	}
	public static int beginOfDay=0;
	public static int endOfDay=0;
	public static int totalDays=0;
	/**
	 * 将excel解析为内存用户对象。
	 * 
	 * @param srcFile
	 * @param beginDay
	 *            开始的列名，含该列
	 * @param endDay
	 *            结束的列名，含该列
	 *            @param isPublic 是不是统计公用座位
	 * @return
	 * @throws BiffException
	 * @throws IOException
	 * @throws ParseException
	 */
	public static Map<String, User> parserExcelToUsers(String srcFile,
			String beginDay, String endDay,boolean isPublic) throws BiffException, IOException,
			ParseException {
		WorkbookSettings setting = new WorkbookSettings();
		java.util.Locale locale = new java.util.Locale("zh", "CN");
		setting.setLocale(locale);
		setting.setEncoding("ISO-8859-1");
		Workbook book = Workbook.getWorkbook(new File(srcFile), setting);
		Sheet sheet = book.getSheet(0);
		Cell cell = null;

		java.text.NumberFormat nf = java.text.NumberFormat.getInstance();
		nf.setMaximumFractionDigits(2);
		// ws.setRowGroup(0, i, false);
		int beginCol = parserColumnName2Int(beginDay);
		int endCol = parserColumnName2Int(endDay);
		cell = sheet.getCell(beginCol, 0);
		String content = cell.getContents();
		beginOfDay=Integer.valueOf(content.substring(0,2));
		cell = sheet.getCell(endCol, 0);
		content = cell.getContents();
		endOfDay=Integer.valueOf(content.substring(0,2));

		totalDays = endCol - beginCol + 1;
		Map<String, User> users = new HashMap<String, User>();
		if (endCol >= sheet.getColumns()) {
			JOptionPane.showMessageDialog(null, "请检查excel格式，列数不对！");
			return null;
		}
		String tmp=null;
		for (int rowIndex = 1; rowIndex < sheet.getRows(); rowIndex++) {// 行
			// 第一列，无意义，可以判断是否到结尾了
			cell = sheet.getCell(0, rowIndex);
			content = cell.getContents();
			if (content.startsWith("记录数:")) {
				break;
			}
			// 第三列，姓名
			cell = sheet.getCell(2, rowIndex);
			String name=null;
			name = cell.getContents();
			if(isPublic){
				name=Utils.publicUsers.get(name);
				if(name==null){
					continue;
				}
			}
			User user = users.get(name);
			if (user == null) {
				// 第二列,人员分类
				cell = sheet.getCell(1, rowIndex);
				user = new User(isPublic?"公用座位统计":cell.getContents(), name, totalDays);
				users.put(name, user);
			}
			// 后续全是时间，一天一列
			for (int d = 0, colIndex = beginCol; colIndex <= endCol; colIndex++, d++) {// 列
				cell = sheet.getCell(colIndex, rowIndex);
				String secondTimeTmp[]=new String[1];
				content = cell.getContents();
				if (content.startsWith("记录数")) {
					break;
				}
				String[] times = content.split("\\s+");
				if(!content.equals("")){
					for(String time:times){
						user.addTime(time, d);
					}
				}
				if(colIndex+1<sheet.getColumns()){
					tmp=sheet.getCell(colIndex+1,rowIndex).getContents();
					if(tmp!=null)
						secondTimeTmp=tmp.split("\\s+");
				}
				if(tmp!=null&&!tmp.equals("")){
					for(String time:secondTimeTmp){
						if(Utils.timeToDouble(time)<=Utils.downTime[3])
							user.addTime(time, d);
					}
				}
				user.fillTime(d);
			}
		}
		return users;
	}


	public static void main(String[] args) throws BiffException, IOException,
	ParseException, RowsExceededException, WriteException {
		JFileChooser chooser=new JFileChooser(new File("E:\\tmp\\考勤"));
		chooser.setVisible(true);
		int returnFile=chooser.showOpenDialog(null);
		if(returnFile!=JFileChooser.APPROVE_OPTION){
			System.exit(1);
		}
		String src= chooser.getSelectedFile().getAbsolutePath();
		String begin=JOptionPane.showInputDialog("请输入开始统计的时间列名（含该列）（区分大小写)","D");
		String end=JOptionPane.showInputDialog("请输入结束统计的时间列名（含该列）（区分大小写)","J");
		try { 
			Map<String,User>users=parserExcelToUsers(src,begin,end,false);
			ArrayList<User>sortUsers=new ArrayList<User>(users.values());
			Collections.sort(sortUsers, new User());
			//重新统计那些公用座位的同学和
			Map<String,User>publicusers=parserExcelToUsers(src,begin,end,true);
			ArrayList<User>sortPublicUsers=new ArrayList<User>(publicusers.values());
			Collections.sort(sortPublicUsers, new User());
			sortUsers.addAll(sortPublicUsers);
			String dest=JOptionPane.showInputDialog("请输入结果文件的保存绝对路径","E:\\tmp\\考勤\\结果1.xls");
			for(User user:sortUsers){
				System.out.println(user);
			}
			writeIntoExcel(dest,sortUsers);
			JOptionPane.showMessageDialog(null, "done!");

			String[] cmd = new String[5];  
			cmd[0] = "cmd";  
			cmd[1] = "/c";  
			cmd[2] = "start";  
			cmd[3] = " ";  
			cmd[4] = new File(dest).getParent();  
			Runtime.getRuntime().exec(cmd);  
		} catch (Exception e) {  
			e.printStackTrace();  
		}  
		System.exit(1);
	}

	private static void writeIntoExcel(String destFile,ArrayList<User> sortUsers) throws FileNotFoundException, IOException, RowsExceededException, WriteException {
		// TODO Auto-generated method stub
		WritableWorkbook wwb = Workbook.createWorkbook(new FileOutputStream(destFile));
		WritableSheet ws1 = wwb.createSheet("考勤统计", 0);
		ws1.addCell(new Label(0, 0, "生成日期："));
		ws1.setColumnView(0, 20);
		insertOneCellData(ws1, 1, 0, new SimpleDateFormat("yyyy-MM-dd HH点").format(new Date()), getDataCellFormat(CellType.DATE));
		ws1.addCell(new Label(0, 1, "分类"));
		ws1.setColumnView(1, 10);
		ws1.addCell(new Label(1, 1, "姓名"));
		ws1.setColumnView(2, 10);
		ws1.addCell(new Label(2, 1, "出勤总时长"));
		ws1.setColumnView(3, 10);
		ws1.addCell(new Label(3, 1, "出勤天数"));
		ws1.setColumnView(4, 10);
		Label labelTitle = new Label(4, 1, "上午");
		setCellComments(labelTitle, "作息规律统计，上午出勤占总出勤的百分比");
		ws1.addCell(labelTitle);
		labelTitle = new Label(5, 1, "下午");
		setCellComments(labelTitle, "作息规律统计，下午出勤占总出勤的百分比");
		ws1.addCell(labelTitle);
		labelTitle = new Label(6, 1, "晚上");
		setCellComments(labelTitle, "作息规律统计，晚上出勤占总出勤的百分比");
		ws1.addCell(labelTitle);
		labelTitle = new Label(7, 1, "夜晚");
		setCellComments(labelTitle, "作息规律统计，加班出勤占总出勤的百分比");
		ws1.addCell(labelTitle);
		ws1.addCell(new Label(8, 1, "备注"));

		WritableSheet ws2 = wwb.createSheet("每日细节", 1);
		ws2.addCell(new Label(0, 1, "分类"));
		ws2.setColumnView(0, 10);
		ws2.addCell(new Label(1, 1, "姓名"));
		ws2.setColumnView(1, 10);


		int realDay=beginOfDay;
		double[][] time=new double[totalDays][5];
		WritableCellFormat  format;
		Colour[] colour=new Colour[]{Colour.RED,Colour.ROSE};
		for(int i=0;i<totalDays;i++){
			format=getDataCellFormat(CellType.LABEL);
			format.setBackground(colour[i%2]);
			insertOneCellData(ws2, i*4+2, 1, realDay+"日上午", format);
			insertOneCellData(ws2, i*4+3, 1, realDay+"日下午", format);
			insertOneCellData(ws2, i*4+4, 1, realDay+"日晚上", format);
			insertOneCellData(ws2, i*4+5, 1, realDay+"日夜", format);
			realDay++;
		}
		int row=2;
		Colour[][] colours=new Colour[][]{{Colour.GRAY_25,Colour.GRAY_50},{Colour.GREY_25_PERCENT,Colour.GREY_40_PERCENT}};
		int level=1;
		for(User user:sortUsers){
			//System.out.println(row+":"+user.getName());
			if(level!=user.getLevel()){
				row++;
				level=user.getLevel();
			}
			insertOneCellData(ws2, 0, row, user.getType(), getDataCellFormat(CellType.STRING_FORMULA));
			insertOneCellData(ws2, 1, row, user.getName(), getDataCellFormat(CellType.STRING_FORMULA));
			time=user.getTime();
			for(int day=0;day<user.getRecords().size();day++){
				//				for(int j=0;j<4;j++){
				//					time[day][j]=user.getLong(day, j);
				//				}
				//				time[day][4]=time[day][0]+time[day][1]+time[day][2]+time[day][3];
				format=getDataCellFormat(CellType.STRING_FORMULA);
				format.setBackground(colours[row%2][day%2]);
				insertOneCellData(ws2,  day*4+2, row, user.toString(day, 0, 0)+"\n"+user.toString(day, 0, 1),  format);
				insertOneCellData(ws2,  day*4+3, row, user.toString(day, 1, 0)+"\n"+user.toString(day, 1, 1),  format);
				insertOneCellData(ws2,  day*4+4, row, user.toString(day, 2, 0)+"\n"+user.toString(day, 2, 1),  format);
				insertOneCellData(ws2,  day*4+5, row, user.toString(day, 3, 0)+"\n"+user.toString(day, 3, 1),  format);
			}
			//			double[] total=new double[]{0,0,0,0,0};
			//			for(int i=0;i<totalDays;i++){
			//				for(int j=0;j<5;j++)
			//					total[j]+=time[i][j];
			//			}
			format=getDataCellFormat(CellType.STRING_FORMULA);format.setBackground(Colour.WHITE);
			insertOneCellData(ws1, 0, row, user.getType(), format);
			format=getDataCellFormat(CellType.STRING_FORMULA);format.setBackground(Colour.LIGHT_ORANGE);
			insertOneCellData(ws1, 1, row, user.getName(), format);
			format=getDataCellFormat(CellType.NUMBER);format.setBackground(Colour.YELLOW);
			insertOneCellData(ws1, 2, row, time[totalDays][4], format);
			format=getDataCellFormat(CellType.NUMBER);format.setBackground(Colour.YELLOW2);
			insertOneCellData(ws1, 3, row, user.getOnlineDays(), format);
			format=getDataCellFormat(CellType.NUMBER);format.setBackground(Colour.GRAY_25);
			insertOneCellData(ws1, 4, row, time[totalDays][4]!=0?time[totalDays][0]/time[totalDays][4]:0, format);
			format=getDataCellFormat(CellType.NUMBER);format.setBackground(Colour.GRAY_50);
			insertOneCellData(ws1, 5, row, time[totalDays][4]!=0?time[totalDays][1]/time[totalDays][4]:0, format);
			format=getDataCellFormat(CellType.NUMBER);format.setBackground(Colour.GRAY_25);
			insertOneCellData(ws1, 6, row, time[totalDays][4]!=0?time[totalDays][2]/time[totalDays][4]:0, format);
			format=getDataCellFormat(CellType.NUMBER);format.setBackground(Colour.GRAY_50);
			insertOneCellData(ws1, 7, row, time[totalDays][4]!=0?time[totalDays][3]/time[totalDays][4]:0, format);
			row++;
		}
		wwb.write();
		wwb.close();

	}
}
