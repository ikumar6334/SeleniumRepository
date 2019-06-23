import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.*;

public class TestFilo {

	public static void main(String[] args) throws FilloException {
		// TODO Auto-generated method stub
		Fillo fillo = new Fillo();
		Connection con = fillo.getConnection("C:\\Users\\admin\\Desktop\\Employee.xlsx");
		String strQuery="Select * from Sheet1 where name='Name1'";
		Recordset recordset=con.executeQuery(strQuery);
		while(recordset.next()){
		System.out.println(recordset.getField("Class"));
		}
		
		String InserQry = "INSERT INTO Sheet2(Name,Country) VALUES('Peter','UK')";
		con.executeUpdate(InserQry);
		recordset.close();
		con.close();
	}

}
