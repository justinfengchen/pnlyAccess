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
	private static String _lymc="�����Ļ���԰"; 
	
	public static void main(String args[]) throws Exception {
		   
		ConnectAccess ca=new ConnectAccess();
		ca.initAccessConnect(); 
		Connection conn=ca.getConnection();
		//����Ӫ������
		//exportCampExcel( conn);
		//����ѧԱ����
		exportStudentExcel( conn);
		
		conn.close();
		
		
	}
	private static void exportStudentExcel(Connection conn) throws Exception
	{
		Map map=new HashMap();
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery("select * from  xyb  where    lymc='"+_lymc+"'");
		FileOutputStream outputStream=new FileOutputStream("d:\\xkb\\"+_lymc+"ѧԱ.xls");
		ExcelUtil excelUtil=new ExcelUtil();
		excelUtil.writeExcel(outputStream);
		List headList= getStudentHeadList01();
		excelUtil.writeHead(headList);
		while (rs.next()) {
			String xyname=rs.getString("xyname") ;//ѧԱ����
			String sex= rs.getString("sex"); // ydbh   Ӫ������
			String birthday=rs.getString("birthday") ; //ydzt Ӫ��״̬

			StringBuilder courseNameSb=new StringBuilder();
			String sbkc1= rs.getString("sbkc1"); //��������
			String sbkc2=rs.getString("sbkc2"); //��԰ 
			String[] skbc1Arr=sbkc1.split(",");
			String[] sbkc2Arr=sbkc2.split(",");
			for(int i=0;i<skbc1Arr.length;i++)
			{
				String ckIdStr=skbc1Arr[i];
				if(ckIdStr==null||ckIdStr.trim().equals(""))
					continue;
				String kcmc=getCCMC( Integer.parseInt(ckIdStr) , conn);
				if("".equals(kcmc))
					throw new Exception("û�пγ̣� "+xyname);
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
					throw new Exception("û�пγ̣� "+xyname);
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
		//����Ӫ�ӱ�Ż�ȡ�ν�
		rs.close();
		stmt.close();
	}
	private static void  exportCampExcel(Connection conn) throws Exception
	{
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery("select * from  ydb  where   ydzt<>'��Ӫ' and lymc='"+_lymc+"'");
		String campName="";
		String campBh="";
		String ydzt="";
		String kcmc="";
		String kkrq="";//��������
		String lymc="";//��������
		FileOutputStream outputStream=new FileOutputStream("d:\\temp\\1.xls");
		ExcelUtil excelUtil=new ExcelUtil();
		excelUtil.writeExcel(outputStream);
		List headList= getHeadList();
		excelUtil.writeHead(headList);
		while (rs.next()) {
			campBh=rs.getString("ydbh") ;//Ӫ�ӱ��  ydbh
			campName= rs.getString("ydmc"); // ydbh   Ӫ������
			ydzt=rs.getString("ydzt") ; //ydzt Ӫ��״̬
			kkrq= rs.getString("kysj"); //��������
			lymc=rs.getString("lymc"); //��԰ 
			String kcId=rs.getString("kcmc"); //�γ�����
			kcmc=getCCMC( Integer.parseInt(kcId) , conn);
			if("".equals(kcmc))
				throw new Exception("û�пγ̣� "+campName);
			List<List> datasList=new ArrayList<List>();
			List rowList=new ArrayList();
			rowList.add(campName);
			rowList.add(campBh);
			rowList.add(ydzt);
			rowList.add(kcmc);
			rowList.add("");//�γ̱��
			rowList.add(kkrq);
			rowList.add("");//�γ̱��
			rowList.add(lymc);
			rowList.add("�ܹ���Ա");
			rowList.add(kkrq);//�ÿ����������
			rowList.add(""); 
			rowList.add(""); 
			rowList.add(""); 
			rowList.add(""); 
			datasList.add(rowList);
			excelUtil.writeData(datasList);
			//��ӿ���
			excelUtil.writeHead(getEmptyList());
			//��ӿνڱ�ͷ 
			excelUtil.writeHead(getLessonHeadList());
			processKj(excelUtil,  conn, campBh);
		}
		excelUtil.close();
		//����Ӫ�ӱ�Ż�ȡ�ν�
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
			String kjs=rs.getString("kjs") ;//�νڱ��
			String kjzt=rs.getString("kczt") ;//�γ�״̬     ��Ҫͳ��һ��
			if("δ��ʼ".equals(kjzt))
				kjzt="�ȴ��Ͽ�";
			else if ("�ѽ���".equals(kjzt))
				kjzt="����";
			else
				throw new Exception("״̬���벻ȫ��"+kjzt);

			if("ͣ��".equals(kjzt))
				throw new Exception("��ͣ�ονڣ�");
			//Ĭ�Ͽν�����ȫ��������
			//��ѧ��ʦ
			String ydzy=rs.getString("ydzy") ; 
			//Ӫ�ӷ�����ʦ
			String ydfy=rs.getString("ydfy") ; 
			//Ӫ�Ӹ�Ӫ��ʦ
			String ydfy1=rs.getString("ydfy1") ; 
			//���Ҵ�Ӫ���д�����
			//�Ͽ�����
			String skrq=rs.getString("pksj") ; 
			String sksj=rs.getString("ydtime") ; 
			String xksj=rs.getString("ydjtime") ; 
			String ydjsId=rs.getString("ydjs");
			String jsmc=getJS(Integer.parseInt(ydjsId), conn);
			if("".equals(jsmc))
				throw new Exception("û�н��ң� Ӫ�ӱ�ţ�"+campBh+"���νڱ�ţ�" +kjs);
			List rowList=new ArrayList();
			rowList.add("");
			rowList.add(kjs);
			rowList.add(kjzt);
			rowList.add("����");//�ν����ͣ�Ĭ������
			rowList.add(ydzy);
			rowList.add(ydfy); 
			rowList.add(ydfy1);
			rowList.add(jsmc);
			rowList.add(skrq); 
			rowList.add(sksj);
			rowList.add(xksj); 
			datasList.add(rowList);
			//����ÿν��µ�ѧԱ
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
				studentRowList.add("����");
				studentRowList.add("");//ѧԱ���
				Integer studentId=Integer.parseInt(ydxyId.trim());
				String studentName=getStudentName(studentId,conn);
				if("".equals(studentName))
					throw new Exception("û��ѧԱ�� Ӫ�ӱ�ţ�"+campBh+"���νڱ�ţ�" +kjs+" ѧԱid��"+studentId);
				studentRowList.add(studentName);
				studentRowList.add("����");
				if(checkIfKQ(studentId,ydkqArr))
				{
					studentRowList.add("����");//��������
					studentRowList.add("�ܹ���Ա");//������
					studentRowList.add("��Ч");//������
					studentRowList.add("��Ч");//������
				}
				else
				{
					studentRowList.add("");//��������
					studentRowList.add("");//������
					studentRowList.add("");//������
					studentRowList.add("");//������
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
				studentRowList.add("����");
				studentRowList.add("");//ѧԱ���
				Integer studentId=Integer.parseInt(ydxyId.trim());
				String studentName=getStudentName(studentId,conn);
				if("".equals(studentName))
					throw new Exception("û��ѧԱ�� Ӫ�ӱ�ţ�"+campBh+"���νڱ�ţ�" +kjs+" ѧԱid��"+studentId);
				studentRowList.add(studentName);
				studentRowList.add("����");   //ѧԱ����
				studentRowList.add("����");   //��������
				studentRowList.add("�ܹ���Ա");//������
				studentRowList.add("��Ч");   //����״̬
				studentRowList.add("��Ч");   //����˵��
				studentDatasList.add(studentRowList);
			}
		} 
		excelUtil.writeData(datasList);

		//��ӿ���
		excelUtil.writeHead(getEmptyList());
		//��� ��ͷ 
		excelUtil.writeHead(getStudentHeadList());
		excelUtil.writeData(studentDatasList);
		//��ӿ���
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
	}	//����list
	private  static List getEmptyList()
	{
		List headList=new ArrayList();
		headList.add("����");
		headList.add("����");
		headList.add("����");
		headList.add("����");
		headList.add("����");
		headList.add("����");
		headList.add("����");
		headList.add("����");
		headList.add("����");
		headList.add("����");
		headList.add("����");
		return headList;
	}
	private static  List getLessonHeadList()
	{
		List headList=new ArrayList();
		headList.add("�ν���Ϣ");
		headList.add("�νڱ��");
		headList.add("�ν�״̬");
		headList.add("�ν�����");
		headList.add("��ѧ��ʦ");
		headList.add("������ʦ");
		headList.add("��Ӫ��ʦ");
		headList.add("���ұ��");
		headList.add("�Ͽ�����");
		headList.add("�Ͽ�ʱ��");
		headList.add("�¿�ʱ��");
		return headList;
	}
	private static  List getStudentHeadList()
	{
		List headList=new ArrayList();
		headList.add("ѧԱ��Ϣ");
		headList.add("�νڱ��");
		headList.add("�ν�����");
		headList.add("ѧԱ���");
		headList.add("ѧԱ����");
		headList.add("ѧԱ����");
		headList.add("��������");
		headList.add("��������");
		headList.add("������");
		headList.add("����״̬");
		headList.add("˵��");
		return headList;
	}
	private static  List getStudentHeadList01()
	{
		List headList=new ArrayList();
		headList.add("ѧԱ����");
		headList.add("�Ա�");
		headList.add("����");
		headList.add("�γ�"); 
		return headList;
	}
	private static  List getHeadList()
	{
		List headList=new ArrayList();
		headList.add("Ӫ������");
		headList.add("Ӫ�ӱ��");
		headList.add("Ӫ��״̬");
		headList.add("�γ�����");
		headList.add("�γ̱��");
		headList.add("��������");
		headList.add("����");
		headList.add("��������");
		headList.add("������");
		headList.add("��������");
		headList.add("��Ӫ�����");
		headList.add("��Ӫ�������");
		headList.add("��Ӫ�������");
		headList.add("��Ӫ����");
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

