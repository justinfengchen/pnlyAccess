import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
public class ConnectAccess {
	private String _dbur1 = "";
	private static String _lymc="常州文化馆园"; 
	
	public static void main(String args[]) throws Exception {
		   
		ConnectAccess ca=new ConnectAccess();
		ca.initAccessConnect(); 
		Connection conn=ca.getConnection();
		//导出营队数据
		//exportCampExcel( conn);
		//导出学员数据
		exportStudentExcel( conn);
		
		conn.close();
		
		
	}
	private static void exportStudentExcel(Connection conn) throws Exception
	{
		Map map=new HashMap();
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery("select * from  xyb  where    lymc='"+_lymc+"'");
		FileOutputStream outputStream=new FileOutputStream("d:\\xkb\\"+_lymc+"学员.xls");
		ExcelUtil excelUtil=new ExcelUtil();
		excelUtil.writeExcel(outputStream);
		List headList= getStudentHeadList01();
		excelUtil.writeHead(headList);
		while (rs.next()) {
			String xyname=rs.getString("xyname") ;//学员名称
			String sex= rs.getString("sex"); // ydbh   营队名称
			String birthday=rs.getString("birthday") ; //ydzt 营队状态

			StringBuilder courseNameSb=new StringBuilder();
			String sbkc1= rs.getString("sbkc1"); //开课日期
			String sbkc2=rs.getString("sbkc2"); //乐园 
			String[] skbc1Arr=sbkc1.split(",");
			String[] sbkc2Arr=sbkc2.split(",");
			for(int i=0;i<skbc1Arr.length;i++)
			{
				String ckIdStr=skbc1Arr[i];
				if(ckIdStr==null||ckIdStr.trim().equals(""))
					continue;
				String kcmc=getCCMC( Integer.parseInt(ckIdStr) , conn);
				if("".equals(kcmc))
					throw new Exception("没有课程！ "+xyname);
				courseNameSb.append(kcmc);
				map.put(kcmc, kcmc);
			}
			for(int i=0;i<sbkc2Arr.length;i++)
			{
				String ckIdStr=sbkc2Arr[i];
				if(ckIdStr==null||ckIdStr.trim().equals(""))
					continue;
				String kcmc=getCCMC( Integer.parseInt(ckIdStr) , conn);
				if("".equals(kcmc))
					throw new Exception("没有课程！ "+xyname);
				courseNameSb.append(kcmc);
				map.put(kcmc, kcmc);
			}
			List<List> datasList=new ArrayList<List>();
			List rowList=new ArrayList();
			rowList.add(xyname);
			rowList.add(sex);
			rowList.add(birthday);
			rowList.add(courseNameSb.toString());// 
			datasList.add(rowList);
			excelUtil.writeData(datasList);
		}
		excelUtil.close();
		//根据营队编号获取课节
		rs.close();
		stmt.close();
	}
	private static void  exportCampExcel(Connection conn) throws Exception
	{
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery("select * from  ydb  where   ydzt<>'结营' and lymc='"+_lymc+"'");
		String campName="";
		String campBh="";
		String ydzt="";
		String kcmc="";
		String kkrq="";//开课日期
		String lymc="";//开课日期
		FileOutputStream outputStream=new FileOutputStream("d:\\temp\\1.xls");
		ExcelUtil excelUtil=new ExcelUtil();
		excelUtil.writeExcel(outputStream);
		List headList= getHeadList();
		excelUtil.writeHead(headList);
		while (rs.next()) {
			campBh=rs.getString("ydbh") ;//营队编号  ydbh
			campName= rs.getString("ydmc"); // ydbh   营队名称
			ydzt=rs.getString("ydzt") ; //ydzt 营队状态
			kkrq= rs.getString("kysj"); //开课日期
			lymc=rs.getString("lymc"); //乐园 
			String kcId=rs.getString("kcmc"); //课程名称
			kcmc=getCCMC( Integer.parseInt(kcId) , conn);
			if("".equals(kcmc))
				throw new Exception("没有课程！ "+campName);
			List<List> datasList=new ArrayList<List>();
			List rowList=new ArrayList();
			rowList.add(campName);
			rowList.add(campBh);
			rowList.add(ydzt);
			rowList.add(kcmc);
			rowList.add("");//课程编号
			rowList.add(kkrq);
			rowList.add("");//课程编号
			rowList.add(lymc);
			rowList.add("总管理员");
			rowList.add(kkrq);//用开课日期替代
			rowList.add(""); 
			rowList.add(""); 
			rowList.add(""); 
			rowList.add(""); 
			datasList.add(rowList);
			excelUtil.writeData(datasList);
			//添加空行
			excelUtil.writeHead(getEmptyList());
			//添加课节表头 
			excelUtil.writeHead(getLessonHeadList());
			processKj(excelUtil,  conn, campBh);
		}
		excelUtil.close();
		//根据营队编号获取课节
		rs.close();
		stmt.close();
	}
	private static void processKj(ExcelUtil excelUtil,Connection conn,String campBh) throws Exception
	{
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery("select * from  skb where   ydbh='"+campBh+"'");
		List<List> datasList=new ArrayList<List>();
		List<List> studentDatasList=new ArrayList<List>();
		while (rs.next()) {
			String kjs=rs.getString("kjs") ;//课节编号
			String kjzt=rs.getString("kczt") ;//课程状态     需要统计一下
			if("未开始".equals(kjzt))
				kjzt="等待上课";
			else if ("已结束".equals(kjzt))
				kjzt="结束";
			else
				throw new Exception("状态翻译不全："+kjzt);

			if("停课".equals(kjzt))
				throw new Exception("有停课课节！");
			//默认课节类型全部是正常
			//教学老师
			String ydzy=rs.getString("ydzy") ; 
			//营队服务老师
			String ydfy=rs.getString("ydfy") ; 
			//营队辅营老师
			String ydfy1=rs.getString("ydfy1") ; 
			//教室从营队中带过来
			//上课日期
			String skrq=rs.getString("pksj") ; 
			String sksj=rs.getString("ydtime") ; 
			String xksj=rs.getString("ydjtime") ; 
			String ydjsId=rs.getString("ydjs");
			String jsmc=getJS(Integer.parseInt(ydjsId), conn);
			if("".equals(jsmc))
				throw new Exception("没有教室！ 营队编号："+campBh+"；课节编号：" +kjs);
			List rowList=new ArrayList();
			rowList.add("");
			rowList.add(kjs);
			rowList.add(kjzt);
			rowList.add("正常");//课节类型，默认正常
			rowList.add(ydzy);
			rowList.add(ydfy); 
			rowList.add(ydfy1);
			rowList.add(jsmc);
			rowList.add(skrq); 
			rowList.add(sksj);
			rowList.add(xksj); 
			datasList.add(rowList);
			//处理该课节下的学员
			String ydxys=rs.getString("ydxy");
			String[] ydxyArr=ydxys.split(",");

			String ydkqs=rs.getString("ydkq");
		    if(ydkqs==null)
		    	ydkqs="";
			String[] ydkqArr=ydkqs.split(",");
			for(int k=0;k<ydxyArr.length;k++)
			{
				String ydxyId=ydxyArr[k];
				if(ydxyId.equals(""))
					continue;
				List studentRowList=new ArrayList();
				studentRowList.add("");
				studentRowList.add(kjs);
				studentRowList.add("正常");
				studentRowList.add("");//学员编号
				Integer studentId=Integer.parseInt(ydxyId.trim());
				String studentName=getStudentName(studentId,conn);
				if("".equals(studentName))
					throw new Exception("没有学员！ 营队编号："+campBh+"；课节编号：" +kjs+" 学员id："+studentId);
				studentRowList.add(studentName);
				studentRowList.add("正常");
				if(checkIfKQ(studentId,ydkqArr))
				{
					studentRowList.add("正常");//考勤类型
					studentRowList.add("总管理员");//考勤人
					studentRowList.add("有效");//考勤人
					studentRowList.add("有效");//考勤人
				}
				else
				{
					studentRowList.add("");//考勤类型
					studentRowList.add("");//考勤人
					studentRowList.add("");//考勤人
					studentRowList.add("");//考勤人
				}
				studentDatasList.add(studentRowList);
			}

			String ydbkkqs=rs.getString("ydbkkq");
			if(ydbkkqs==null)
				ydbkkqs="";
			String[] ydbkkqArr=ydbkkqs.split(",");
			for(int k=0;k<ydbkkqArr.length;k++)
			{
				String ydxyId=ydbkkqArr[k];
				if(ydxyId.equals(""))
					continue;
				List studentRowList=new ArrayList();
				studentRowList.add("");
				studentRowList.add(kjs);
				studentRowList.add("正常");
				studentRowList.add("");//学员编号
				Integer studentId=Integer.parseInt(ydxyId.trim());
				String studentName=getStudentName(studentId,conn);
				if("".equals(studentName))
					throw new Exception("没有学员！ 营队编号："+campBh+"；课节编号：" +kjs+" 学员id："+studentId);
				studentRowList.add(studentName);
				studentRowList.add("正常");   //学员类型
				studentRowList.add("正常");   //考勤类型
				studentRowList.add("总管理员");//考勤人
				studentRowList.add("有效");   //考勤状态
				studentRowList.add("有效");   //考勤说明
				studentDatasList.add(studentRowList);
			}
		} 
		excelUtil.writeData(datasList);

		//添加空行
		excelUtil.writeHead(getEmptyList());
		//添加 表头 
		excelUtil.writeHead(getStudentHeadList());
		excelUtil.writeData(studentDatasList);
		//添加空行
		excelUtil.writeHead(getEmptyList()); 
		rs.close();
		stmt.close();
	}
	

	
	/**
	 * 
	 * @param id
	 * @param ydkqArr
	 * @return
	 */
	private static boolean checkIfKQ(Integer id,String[] ydkqArr)
	{
		for(int i=0;i<ydkqArr.length;i++)
		{
			String ydxyId=ydkqArr[i];
			if(ydxyId.equals(""))
				continue;
			Integer tempId=Integer.parseInt(ydxyId.trim());
			if(tempId.intValue()==id.intValue())
			{
				return true;
			}
		}
		return false;
	}
	private static String getStudentName(Integer id,Connection conn) throws SQLException
	{
		String sql="Select xyname from xyb where id="+id;
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery(sql);
		String xyname="";
		while (rs.next())
		{
			xyname=rs.getString("xyname");
			break;
		}
		rs.close();
		stmt.close();
		return xyname;
	}
	public void initAccessConnect() throws Exception 
	{
		Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
		/**
		 */
		//_dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=D://workspace//2015//xkb.mdb";
	
		_dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=D://xkb//xkb.mdb";
		
		
	}
	public Connection getConnection() throws SQLException
	{
		Connection conn = DriverManager.getConnection(_dbur1, "", "");
		return conn;
	}	//空行list
	private  static List getEmptyList()
	{
		List headList=new ArrayList();
		headList.add("空行");
		headList.add("空行");
		headList.add("空行");
		headList.add("空行");
		headList.add("空行");
		headList.add("空行");
		headList.add("空行");
		headList.add("空行");
		headList.add("空行");
		headList.add("空行");
		headList.add("空行");
		return headList;
	}
	private static  List getLessonHeadList()
	{
		List headList=new ArrayList();
		headList.add("课节信息");
		headList.add("课节编号");
		headList.add("课节状态");
		headList.add("课节类型");
		headList.add("教学老师");
		headList.add("服务老师");
		headList.add("辅营老师");
		headList.add("教室编号");
		headList.add("上课日期");
		headList.add("上课时间");
		headList.add("下课时间");
		return headList;
	}
	private static  List getStudentHeadList()
	{
		List headList=new ArrayList();
		headList.add("学员信息");
		headList.add("课节编号");
		headList.add("课节类型");
		headList.add("学员编号");
		headList.add("学员名称");
		headList.add("学员类型");
		headList.add("考勤类型");
		headList.add("考勤类型");
		headList.add("考勤人");
		headList.add("考勤状态");
		headList.add("说明");
		return headList;
	}
	private static  List getStudentHeadList01()
	{
		List headList=new ArrayList();
		headList.add("学员名称");
		headList.add("性别");
		headList.add("生日");
		headList.add("课程"); 
		return headList;
	}
	private static  List getHeadList()
	{
		List headList=new ArrayList();
		headList.add("营队名称");
		headList.add("营队编号");
		headList.add("营队状态");
		headList.add("课程名称");
		headList.add("课程编号");
		headList.add("开课日期");
		headList.add("描述");
		headList.add("所属部门");
		headList.add("创建人");
		headList.add("创建日期");
		headList.add("开营审核人");
		headList.add("开营审核日期");
		headList.add("结营队审核人");
		headList.add("结营日期");
		return headList;
	}private static String  getCCMC(Integer kcId,Connection conn) throws SQLException
	{
		String sql="select * from kcmcb where  id="+kcId;
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery(sql);
		String kcmc="";
		while (rs.next())
		{
			kcmc=rs.getString("kcmc");
		}
		rs.close();
		stmt.close();
		return kcmc;
	}
	private static String  getJS(Integer jsId,Connection conn) throws SQLException
	{
		String sql="select * from jcb where  id="+jsId;
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery(sql);
		String kcmc="";
		while (rs.next())
		{
			kcmc=rs.getString("jcmc");
			break;
		}
		rs.close();
		stmt.close();
		return kcmc;
	}
}

