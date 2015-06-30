

import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


import jxl.SheetSettings;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelUtil extends Thread {
	/** 空方法，供初始化 */
	public ExcelUtil() {
	}

	// =============================== 方法控制区 ==============================

 

	/**
	 * 写入Excel流
	 * 
	 * @param os
	 * @param style
	 */
	public void writeExcel(OutputStream os) {
		initExcel(os);
	}

	/**
	 * 初始化Excel流
	 * @param filePath
	 * @return
	 */
	private WritableSheet initExcel(OutputStream os) {
		try {
			this.os = os;
			wwb = Workbook.createWorkbook(os);
			ws = wwb.createSheet("sheet1", 0); // 建立工作表			SheetSettings ss = ws.getSettings();
			ss.setFitHeight(dataHeight);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return ws;
	}

	/**
	 * 写表头
	 * 
	 * @param ws
	 * @throws WriteException
	 * @throws RowsExceededException
	 */
	public void writeHead(List headList) throws RowsExceededException, WriteException {
		if (headList == null || headList.size()==0)  
			return;
		for (int j = 0  ; j < headList.size(); j++  )
		{
			WritableCellFormat wf = getStyle(2);
			ws.addCell(new Label(j, startRowNum, (String) headList.get(j), wf));
			System.out.println(startRowNum +" "+ (String) headList.get(j) );
			int colLength = getWidth((String) headList.get(j));
			if (ws.getColumnView(j).getSize() < (colLength + fixWidth) * 256) 
			{
				ws.setColumnView(j, colLength + fixWidth);
			}
		}
		startRowNum++;
	}
	/**
	 * 写Excel
	 * 
	 * @param filePath
	 * @param map
	 * @return
	 * @throws Exception
	 */
	public void close( ) {
		try {
			wwb.write(); // 确定写入
		} catch (Exception e) {
			System.out.println("error in ExcelUtil writeExcel()");
			e.printStackTrace();
		} finally {
			try {
				if (null != wwb) {
					wwb.close();
				}
				if (null != os) {
					os.close();
				}
			} catch (Exception e) {
				System.out.println("error in ExcelUtil writeExcel()");
				e.printStackTrace();
			}
		}
	}
	/**
	 * 写入数据区
	 * 
	 * 
	 * @param ws
	 * @throws WriteException
	 */
	@SuppressWarnings("unchecked")
	public void writeData(List dataList) throws WriteException {
		if (dataList== null || dataList.size()==0) 
			return;
		for (int i = 0; i < dataList.size(); i++, startRowNum++) {
			List row = (List) dataList.get(i);
			for (int j = 0, field = 0; field < row.size(); j++, field++) {
				if (row.get(field) instanceof String) {
					WritableCellFormat wf = getStyle(2);
					ws.addCell(new Label(j, startRowNum, (String) row
							.get(field), wf));
					int colLength = getWidth((String) row.get(field));
					if (ws.getColumnView(j).getSize() < (colLength + fixWidth) * 256) {
						ws.setColumnView(j, colLength + fixWidth);
					}
				} else if (row.get(field) instanceof Integer) {
					WritableCellFormat wf = null;
					if (amtDisplay.indexOf("" + field) != -1) {
						wf = getStyle(31);
					} else {
						wf = getStyle(3);
					}
					ws.addCell(new Number(j, startRowNum, ((Integer) row
							.get(field)).intValue(), wf));
				} else if (row.get(field) instanceof Double) {
					WritableCellFormat wf = null;
					if (amtDisplay.indexOf("" + field) != -1) {
						wf = getStyle(31);
					} else {
						wf = getStyle(3);
					}
					ws.addCell(new Number(j, startRowNum, ((Double) row
							.get(field)).doubleValue(), wf));
				} else if (row.get(field) instanceof BigDecimal) {
					WritableCellFormat wf = null;
					wf = getStyle(31);
					ws.addCell(new Number(j, startRowNum, ((BigDecimal) row
							.get(field)).doubleValue(), wf));
					int colLength = getWidth(row.get(field).toString());
					if (ws.getColumnView(j).getSize() < (colLength + fixWidth) * 256) {
						ws.setColumnView(j, colLength + fixWidth);
					}
				} else if (row.get(field) == null) {
					WritableCellFormat wf = getStyle(2);
					ws.addCell(new Label(j, startRowNum, "", wf));
				} else {
					WritableCellFormat wf = getStyle(2);
					ws.addCell(new Label(j, startRowNum, String.valueOf(row
							.get(field)), wf));
				}
			}
		}
	}


	/**
	 * 合并单元格算法
	 * 
	 * 
	 * @param ws
	 */
	public void span(WritableSheet ws) {
		if (spanKey != -1) {
			for (int i = 0; i < ws.getRows(); i++) {
				int start = i;
				String tmpStr = ws.getCell(spanKey, i).getContents();
				for (int j = i + 1; j <= ws.getRows(); j++) {
					if (j < ws.getRows()
							&& tmpStr.equals(ws.getCell(spanKey, j)
									.getContents())) {
						continue;
					} else if (j - i == 1) {
						break;
					} else {
						try {
							ws.mergeCells(spanKey, i, spanKey, j - 1);
						} catch (Exception e) {
							System.out.println("error in ExcelUtil span()");
							e.printStackTrace();
						}
						i = j - 1;
						break;
					}
				}
				for (int n = 0; n < spanRows.size(); n++) {
					int col = Integer.parseInt((String) spanRows.get(n));
					if (col == spanKey) {
						continue;
					}
					for (int k = start; k <= i; k++) {
						tmpStr = ws.getCell(col, k).getContents();
						for (int m = k + 1; m <= i + 1; m++) {
							if (m < i + 1
									&& tmpStr.equals(ws.getCell(col, m)
											.getContents())) {
								continue;
							} else if (m - k == 1) {
								break;
							} else {
								try {
									ws.mergeCells(col, k, col, m - 1);
								} catch (Exception e) {
									System.out
											.println("error in ExcelUtil span()");
									e.printStackTrace();
								}
								k = m - 1;
								break;
							}
						}
					}
				}
			}
		}
	}

	/**
	 * 设置边框，字体，对齐方式 0 - 水平垂直居中加粗 1 - 水平垂直居中 2 - 水平居左垂直居中 3 - 水平居右垂直居中 31 -
	 * 3的金额方式(#0.00)显示 4 - 水平居左垂直居中 ， 显示上左下边框
	 * 
	 * 5 - 水平居右垂直居中 ， 显示上右下边框
	 * 
	 * @param pos
	 * @return
	 */
	public WritableCellFormat getStyle(int pos) {
		if (wf == null) {
			wf = new HashMap<Integer, WritableCellFormat>();
			try {
				WritableCellFormat wf0 = new WritableCellFormat();
				wf0.setBorder(Border.ALL, BorderLineStyle.THIN);
				wf0.setAlignment(CENTER);
				wf0.setVerticalAlignment(MIDDLE);
				wf0.setFont(new WritableFont(WritableFont.createFont("Arial"),
						14, WritableFont.BOLD));
				wf0.setWrap(true);
				wf.put(0, wf0);
				WritableCellFormat wf1 = new WritableCellFormat();
				wf1 = new WritableCellFormat();
				wf1.setBorder(Border.ALL, BorderLineStyle.THIN);
				wf1.setAlignment(CENTER);
				wf1.setVerticalAlignment(MIDDLE);
				wf1.setWrap(true);
				wf.put(1, wf1);
				WritableCellFormat wf2 = new WritableCellFormat();
				wf2 = new WritableCellFormat();
				wf2.setBorder(Border.ALL, BorderLineStyle.THIN);
				wf2.setAlignment(LEFT);
				wf2.setVerticalAlignment(MIDDLE);
				wf2.setWrap(true);
				wf.put(2, wf2);
				WritableCellFormat wf3 = new WritableCellFormat();
				wf3 = new WritableCellFormat();
				wf3.setBorder(Border.ALL, BorderLineStyle.THIN);
				wf3.setAlignment(RIGHT);
				wf3.setVerticalAlignment(MIDDLE);
				wf3.setWrap(true);
				wf.put(3, wf3);
				WritableCellFormat wf31 = new WritableCellFormat();
				wf31 = new WritableCellFormat(new NumberFormat("#0.00"));
				wf31.setBorder(Border.ALL, BorderLineStyle.THIN);
				wf31.setAlignment(RIGHT);
				wf31.setVerticalAlignment(MIDDLE);
				wf31.setWrap(true);
				wf.put(31, wf31);
				WritableCellFormat wf4 = new WritableCellFormat();
				wf4.setBorder(Border.TOP, BorderLineStyle.THIN);
				wf4.setBorder(Border.LEFT, BorderLineStyle.THIN);
				wf4.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
				wf4.setAlignment(LEFT);
				wf4.setVerticalAlignment(MIDDLE);
				wf4.setWrap(true);
				wf.put(4, wf4);
				WritableCellFormat wf5 = new WritableCellFormat();
				wf5.setBorder(Border.TOP, BorderLineStyle.THIN);
				wf5.setBorder(Border.RIGHT, BorderLineStyle.THIN);
				wf5.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
				wf5.setAlignment(RIGHT);
				wf5.setVerticalAlignment(MIDDLE);
				wf5.setWrap(true);
				wf.put(5, wf5);
			} catch (WriteException e) {
				System.out.println("error in ExcelUtil getStyle()");
				e.printStackTrace();
			}
		}

		return wf.get(pos);
	}
  
	// =============================== 工具方法区==============================

	/**
	 * 字符串和集合转换
	 * 
	 * @param s
	 * @param vec
	 */
	@SuppressWarnings("unchecked")
	private void stringToList(String s, List vec) {
		if (s != null && !s.equals("")) {
			String[] ss = s.split("\\,");
			vec.clear();
			for (int i = 0; i < ss.length; i++) {
				vec.add(ss[i]);
			}
		}
	}

	/**
	 * 取字符串宽度，字母数字宽度1，汉字宽度2，其余宽度1
	 * 
	 * @param s
	 * @return
	 */
	public int getWidth(String s) {
		if (s == null || s.length() == 0) {
			return 0;
		}
		int len = 0;
		for (int i = 0; i < s.length(); i++) {
			char c = s.charAt(i);
			if (isLetter(c)) {
				len++;
			} else {
				if (Character.isLetter(c)) {
					len += 2;
				} else {
					len++;
				}
			}
		}
		if (len > 50)
			len = 50;
		return len;
	}

	/**
	 * 判断一个字符是不是单字节字符
	 * 
	 * 
	 * @param char c
	 * @return boolean
	 */
	public static boolean isLetter(char c) {
		int k = 0x80;
		return c / k == 0;
	}

	// =============================== 常量定义区==============================

	public static final Alignment CENTER = Alignment.CENTRE;
	public static final Alignment LEFT = Alignment.LEFT;
	public static final Alignment RIGHT = Alignment.RIGHT;
	public static final VerticalAlignment TOP = VerticalAlignment.TOP;
	public static final VerticalAlignment MIDDLE = VerticalAlignment.CENTRE;
	public static final VerticalAlignment BOTTOM = VerticalAlignment.BOTTOM;

	// =============================== 样式变量区==============================

	/** 标题行高 */
	public int titleHeight = 400;

	/** 数据区行高 */
	public int dataHeight = 300;

	/** 默认列宽 */
	public int defaultWidth = 8;
	/**
	 * <pre>
	 * 修正列宽，默认列宽是根据10号字体大小的， 如果将字符放大，
	 * 需要多加一些修正列宽， 以免显示不漂亮，这里值代表增加多少个单字节字的宽度
	 * 
	 * </pre>
	 */
	public int fixWidth = 2;

	// =============================== 数据变量区==============================

	/** Excel标题头 */
	private String head = "";

	/** 左子表头 */
	private String left_head = "";

	/** 右子表头 */
	private String right_head = "";

	/** 中间子表头 */
	private String middle_head = "";

	/** 标题栏，可以为空 */
	@SuppressWarnings("unchecked")
	private List title = new ArrayList();

	/** 标题栏，可以为空 如果设置了titleStr忽视title属性 多值以逗号分隔 */
	private String titleStr = "";

	private int spanKey = -1;
	@SuppressWarnings("unchecked")
	private List spanRows = new ArrayList();
	private String spanRowsStr = "";

	// =============================== 隐藏设置变量 ==============================

	// =============================== 统计行设置变量 ==============================

	private String totalStr = "";
	@SuppressWarnings("unchecked")
	private List total = new ArrayList();
	@SuppressWarnings("unchecked")
	private List totalRet = new ArrayList();

	// =============================== 金额方式显示变量 ==============================

	private String amtDisplayStr = "";
	@SuppressWarnings("unchecked")
	private List amtDisplay = new ArrayList();

	// =============================== 统计行样式变量 ==============================

	/** 统计行边框显示方式 1默认 所有单元格套用样式 0只有有数据的单元格才套用样式 */
	public int totalRowBorder = 1;
	/**
	 * <pre>
	 * 统计行边框显示样式	 * 
	 * 1默认，只处理左右下边框	 * 
	 * 2 处理所有边框
	 * 
	 * </pre>
	 */
	// public int totalBorderPos = 1;

	/**
	 * <pre>
	 * 统计行显示样式	 * 
	 * 1默认，合并最后一行，所有数据放在一个单元格中，每个数据中加入separator对应的字符串
	 * 2不合并，每个数据一个单元格，每个单元格之间空block个单元格
	 * 3不合并，按照totalcols属性规定具体的列来显示统计结果
	 * 4合并，按照totalcols属性规定具体的列来显示统计结果，并根据totalspancols向后合并列
	 * 
	 * </pre>
	 */
	public int totalStyle = 1;

	/** 统计行两个数据中间的分隔符 */
	public String totalseparator = "   ";

	/** 统计行两个单元格之间的空格 */
	public int totalblock = 1;

	/** 显示的统计结果的具体列，多列以逗号分隔 */
	public String totalcols = "";

	/**
	 * <pre>
	 * 每个统计结果合并的列的数目，多列以逗号分隔
	 * 0表示不合并，1表示合并一列，依次类推
	 * 注意：如果合并列过多（末列大于了下一个统计结果的列号），可能导致下一个统计结果不显示
	 * </pre>
	 */
	public String totalspancols = "";

	// =============================== 内部变量 ==============================

	/**
	 * Excel对象
	 */
	private WritableWorkbook wwb = null;
	/**
	 * 输出流
	 */
	private OutputStream os = null;
	/**
	 * Excel工作表
	 */
	public WritableSheet ws = null;
	/**
	 * 起始行
	 */
	public int startRowNum = 0;
	/**
	 * 样式变量
	 */
	private Map<Integer, WritableCellFormat> wf = null;

	// =========================== setter and getter
	// ============================
	/** Excel标题头 */
	public void setHead(String head) {
		this.head = head;
	}

	/** 左子表头 */
	public void setLeft_head(String left_head) {
		this.left_head = left_head;
	}

	/** 右子表头 */
	public void setRight_head(String right_head) {
		this.right_head = right_head;
	}

	/** 中间子表头 */
	public void setMiddle_head(String middleHead) {
		middle_head = middleHead;
	}

	/** 标题栏，可以为空 多值以逗号分隔 */
	public void setTitleStr(String titleStr) {
		this.titleStr = titleStr;
	}
	/** 合并关键字，如果为-1则不进行合并 */
	public void setSpanKey(int spanKey) {
		this.spanKey = spanKey;
	}

	/** 标题栏，可以为空 */
	public void setSpanRowsStr(String spanRowsStr) {
		this.spanRowsStr = spanRowsStr;
	}

  
	public void setTotalStr(String totalStr) {
		this.totalStr = totalStr;
	}

	/**
	 * 设置以金额方式显示的列号，多列以逗号分隔
	 */
	public void setAmtDisplayStr(String str) {
		this.amtDisplayStr = str;
	}

}
