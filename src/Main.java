import java.awt.AWTException;
import java.awt.BorderLayout;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Scanner;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
	public static void main(String[] args) {
		new Base_window();
	}
}	

class CantSave extends JDialog {
	private static final long serialVersionUID = 1510808506579842063L;
	
	JLabel str = new JLabel("저장에 실패했습니다");
	JButton close = new JButton("확인");
	
	CantSave(JFrame content) {
		super(content, "Error");
		str.setFont(new Font("고딕", Font.PLAIN, 25));
		setLayout(new BorderLayout());
		setSize(270, 180);
		setLocation(270, 180);
		setDefaultCloseOperation(DISPOSE_ON_CLOSE);
		
		close.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				setVisible(false);
			}
		});
		
		str.setHorizontalAlignment(JLabel.CENTER);
		
		add(str, "Center");
		add(close, "South");
	}
}

class ComSave extends JDialog {
	private static final long serialVersionUID = 698919998981537952L;
	
	JLabel come = new JLabel("출석 확인되었습니다");
	JLabel goHome = new JLabel("귀가 확인되었습니다");
	JButton close = new JButton("확인");
	
	ComSave(JFrame content, String state) {
		super(content, "Save");
		come.setFont(new Font("고딕", Font.PLAIN, 25));
		goHome.setFont(new Font("고딕", Font.PLAIN, 25));
		setLayout(new BorderLayout());
		setSize(270, 180);
		setLocation(270, 180);
		setDefaultCloseOperation(DISPOSE_ON_CLOSE);
		
		close.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				setVisible(false);
			}
		});
		
		if(state == "출석") {
			come.setHorizontalAlignment(JLabel.CENTER);
			add(come, "Center");
		} else {
			goHome.setHorizontalAlignment(JLabel.CENTER);
			add(goHome, "Center");
		}
		add(close, "South");
	}
}

class Base_window extends JFrame {
	private static final long serialVersionUID = -2233388981697358L;
	
	JLabel attNum = new JLabel("출석 번호");
	JTextField inNum = new JTextField(15);
	JPanel base = new JPanel();
	JPanel lower = new JPanel();
	JButton ok = new JButton("확인");
	CantSave err = new CantSave(this);
	String inputNum = new String();
	File file  = new File("D:\\Workspace\\Attendance\\src\\data.xlsx");
	int fieldrows = howManyPeople();
	boolean[] isThere = new boolean[fieldrows];
	String copyStr = new String();
	ComSave okSave;
	
	Base_window() {
		setTitle("Attendance Management");
		setSize(480,360);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		appInit();
		
		attNum.setFont(new Font("나눔 고딕", Font.BOLD, 60));
		attNum.setHorizontalAlignment(JLabel.CENTER);
		
		inNum.setFont(new Font("고딕",Font.PLAIN, 40));
		inNum.setHorizontalAlignment(JLabel.CENTER);
		
		inNum.addActionListener(saveAction);
		ok.addActionListener(saveAction);
		
		lower.setLayout(new GridLayout(1,2));
		lower.add(inNum);
		lower.add(ok);
		
		base.setLayout(new GridLayout(2,1));
		base.add(attNum);
		base.add(lower);
		
		add(base);
		setVisible(true);
	}
	
	void appInit( ) {
		arrayInit(isThere);
		openMessenger();
	}
	
	ActionListener saveAction = new ActionListener( ) {
		@Override
		public void actionPerformed(ActionEvent e) {
			if(isNumber(inNum.getText())) {
				inputNum = inNum.getText();
				if(!isThere[Integer.parseInt(inputNum)]) {
					inNum.setText("");
					attSave("출석");
					isThere[Integer.parseInt(inputNum)] = true;
					sendMsg("출석");
					typeMsg();
				} else {
					inNum.setText("");
					attSave("귀가");
					isThere[Integer.parseInt(inputNum)] = false;
					sendMsg("귀가");
					typeMsg();
				}
			} else {
				inNum.setText("");
				err.setVisible(true);
			}
		}	
	};
	
	int howManyPeople() {
		int howMany = 0;
		try {
			XSSFWorkbook Wb = new XSSFWorkbook(new FileInputStream(file));
			XSSFSheet fieldSheet = Wb.getSheetAt(1);
			howMany = fieldSheet.getPhysicalNumberOfRows();
			Wb.close();
		}
		catch(FileNotFoundException fe) {
			System.out.println("Not Found Exception >> " + fe.toString());
		}
		catch(IOException ie) {
			System.out.println("IOException >> " + ie.toString());		
		}
		return howMany;
	}
	
	void arrayInit(boolean[] arr) {
		for(int i=0;i<arr.length;i++) {
			arr[i] = false;
		}
	}
	
	void attSave(String state) {
		okSave = new ComSave(this, state);
		
		try {
			XSSFWorkbook xlsxWb = new XSSFWorkbook(new FileInputStream(file));
			XSSFSheet wkSheet = xlsxWb.getSheetAt(0);
			XSSFCellStyle dateStyle = xlsxWb.createCellStyle();
			XSSFCellStyle timeStyle = xlsxWb.createCellStyle();
			XSSFDataFormat dataForm = xlsxWb.createDataFormat();
			Scanner datain = new Scanner(System.in);
			int rows = wkSheet.getPhysicalNumberOfRows();
			int input = 0;
			Cell numCell = null;
			Cell dateCell = null;
			Cell timeCell = null;
			Cell nameCell = null;
			Cell stateCell = null;
			int rowindex = rows + 1;
		
			dateStyle.setDataFormat(dataForm.getFormat("yyyy-mm-dd"));
			timeStyle.setDataFormat(dataForm.getFormat("h:mm:ss;@"));
			
			XSSFRow row = wkSheet.createRow(rows);
			
			input = Integer.parseInt(inputNum);
			
			numCell = row.createCell(0);
			numCell.setCellValue(input); 		//출석 번호 저장
			
			nameCell = row.createCell(1);
			nameCell.setCellFormula("VLOOKUP(A"+ rowindex + ",Sheet2!$A$2:$B$11,2,FALSE)"); //수식 저장
			
			dateCell = row.createCell(2);
			dateCell.setCellValue(Calendar.getInstance());
			dateCell.setCellStyle(dateStyle); 	//출석 날짜 저장
			
			timeCell = row.createCell(3);
			timeCell.setCellValue(Calendar.getInstance()); 
			timeCell.setCellStyle(timeStyle);	//출석 시간 저장
			
			stateCell = row.createCell(4);
			stateCell.setCellValue(state);		//출석 귀가 여부
			
			FileOutputStream fos = new FileOutputStream(file);
			xlsxWb.write(fos);
			datain.close();
			xlsxWb.close();
		}
		catch(FileNotFoundException fe) {
			System.out.println("Not Found Exception >> " + fe.toString());
		}
		catch(IOException ie) {
			System.out.println("IOException >> " + ie.toString());		
		}
		catch(NumberFormatException ne) {
			System.out.println("NumberFormetException >> " + ne.toString());
			err.setVisible(true);
		}
		
		okSave.setVisible(true);
	}
	
	void sendMsg(String remark) {
		
		try {
			XSSFWorkbook xlsxWb = new XSSFWorkbook(new FileInputStream(file));
			XSSFSheet wkSheet = xlsxWb.getSheetAt(0);
			XSSFFormulaEvaluator fev = xlsxWb.getCreationHelper().createFormulaEvaluator();
			DataFormatter df = new DataFormatter();
			int rows = wkSheet.getPhysicalNumberOfRows();
			XSSFRow row = wkSheet.getRow(rows-1);
			XSSFCell name = row.getCell(1);
			
			String sName = df.formatCellValue(name, fev);
			Calendar cld = Calendar.getInstance();
			SimpleDateFormat dateForm = new SimpleDateFormat("yy-MM-dd HH:mm:ss");
			String sDate = dateForm.format(cld.getTime());
			
			copyStr = sName + " " + sDate + " " + remark;
			
			xlsxWb.close();
		}
		catch(FileNotFoundException fe) {
			System.out.println("Not Found Exception >> " + fe.toString());
		}
		catch(IOException ie) {
			System.out.println("IOException >> " + ie.toString());		
		}
		
	}
	
	void typeMsg(){
		Clipboard cb = Toolkit.getDefaultToolkit().getSystemClipboard();
		StringSelection contents = new StringSelection(copyStr);
		Robot macro;
		
		try {
			macro = new Robot();
			
			cb.setContents(contents, null);
			
			macro.mouseMove(1100, 500);
			macro.mousePress(InputEvent.BUTTON1_MASK);
			macro.mouseRelease(InputEvent.BUTTON1_MASK);
			macro.delay(50);
			macro.keyPress(KeyEvent.VK_CONTROL);
			macro.keyPress(KeyEvent.VK_V);
			macro.keyRelease(KeyEvent.VK_CONTROL);
			macro.keyRelease(KeyEvent.VK_V);
			macro.delay(50);
			macro.keyPress(KeyEvent.VK_ENTER);
			macro.keyRelease(KeyEvent.VK_ENTER);
			macro.delay(50);
			macro.mouseMove(300, 300);
			macro.mousePress(InputEvent.BUTTON1_MASK);
			macro.mouseRelease(InputEvent.BUTTON1_MASK);
		}
		catch (AWTException ae) {
			// TODO Auto-generated catch block
			System.out.println("AWTException >> " + ae.toString());
		}
	}
	
	void openMessenger() {
		Process messenger;
		try {
			messenger = new ProcessBuilder("notepad.exe").start();
		}
		catch (IOException ie) {
			System.out.println("IOException >> " + ie.toString());
		}
	}
	
	static boolean isNumber(String str) {
		try {
			Double.parseDouble(str);
			return true;
		}
		catch(NumberFormatException e) {
			return false;
		}
	}
}