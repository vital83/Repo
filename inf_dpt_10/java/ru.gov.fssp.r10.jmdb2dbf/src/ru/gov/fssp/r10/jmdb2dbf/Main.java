package ru.gov.fssp.r10.jmdb2dbf;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import com.linuxense.javadbf.*;




public class Main {
    public static void main(String[] args)
    {
        String txtConStrMdb, txtConStrDbf, txtConStrLog, txtLogFileName;
        String txtLog = "";

        // forName("net.ucanaccess.jdbc.UcanaccessDriver");

        txtConStrMdb = "jdbc:ucanaccess://D:/Repository/java/mdb2dbf/data/PrRoz.mdb";
        txtConStrDbf = "d:\\Repository\\java\\mdb2dbf\\data\\prroz.dbf";
        txtLogFileName = "mdb2dbf.log";

        try {

            // настройки из InFile.txt
            List<String> inParams = ReadFileLinesOld("InFile.txt", StandardCharsets.UTF_8);
            txtConStrMdb = inParams.get(0);
            txtConStrDbf = inParams.get(1);
            txtLogFileName = inParams.get(2);

            Class.forName("net.ucanaccess.jdbc.UcanaccessDriver"); // динамическая загрузка класса
            Connection connection = DriverManager.getConnection(txtConStrMdb);
            Statement statement = connection.createStatement();

            System.out.println("Connected to MS Access database.");
            String sqlQuery = "SELECT nsyst, kodrai_kod, kateg, fam, imj, otch, dt_rojd_d, dt_rojd_m, dt_rojd_y, mes_rojd, nomer, dt_to_gic, mfr, y_uch, m_uch, d_uch, svertka FROM PrRoz";
            try (ResultSet resultSet = statement.executeQuery(sqlQuery)) {
                System.out.println("MSAccess select query executed.");
                int result = writeResultSetToDBF(resultSet, txtConStrDbf);
            }

            connection.close();

            } catch (SQLException e) {
                System.err.println("SQLException.");
                e.printStackTrace();
                //throw new RuntimeException(e);
        } catch (ClassNotFoundException e) {
            System.err.println("ClassNotFoundException.");
            e.printStackTrace();
            //throw new RuntimeException(e);
        }
    }


    /* read dbf example
            Connection con = DriverManager.getConnection( "jdbc:dbschema:dbf:/sample_dbf_folder" );
            Statement st = con.createStatement();
            ResultSet rs = st.executeQuery("select * from cars")
            while( rs.next() ){
                //your code here
            }
     */

    private static DBFField[] prepareFields(){

        DBFField[] fields = new DBFField[17];

        // последовательность полей должна быть такая же как последовательность в sql выборке из MSAccess

        // nsyst	numeric 13
        fields[0] = new DBFField();
        fields[0].setName("nsyst");
        fields[0].setType(DBFDataType.NUMERIC);
        fields[0].setLength(13);

        // kodrai_kod	character 110
        fields[1] = new DBFField();
        fields[1].setName("kodrai_kod");
        fields[1].setType(DBFDataType.CHARACTER);
        fields[1].setLength(110);

        // kateg	character 50
        fields[2] = new DBFField();
        fields[2].setName("kateg");
        fields[2].setType(DBFDataType.CHARACTER);
        fields[2].setLength(50);

        // fam	character 80
        fields[3] = new DBFField();
        fields[3].setName("fam");
        fields[3].setType(DBFDataType.CHARACTER);
        fields[3].setLength(80);

        // imj	character 60
        fields[4] = new DBFField();
        fields[4].setName("imj");
        fields[4].setType(DBFDataType.CHARACTER);
        fields[4].setLength(60);

        // otch	character 60
        fields[5] = new DBFField();
        fields[5].setName("otch");
        fields[5].setType(DBFDataType.CHARACTER);
        fields[5].setLength(60);

        // dt_rojd_d	numeric 20
        fields[6] = new DBFField();
        fields[6].setName("dt_rojd_d");
        fields[6].setType(DBFDataType.NUMERIC);
        fields[6].setLength(20);

        // dt_rojd_m	numeric 20
        fields[7] = new DBFField();
        fields[7].setName("dt_rojd_m");
        fields[7].setType(DBFDataType.NUMERIC);
        fields[7].setLength(20);

        // dt_rojd_y	numeric 20
        fields[8] = new DBFField();
        fields[8].setName("dt_rojd_y");
        fields[8].setType(DBFDataType.NUMERIC);
        fields[8].setLength(20);

        // mes_rojd	character 250
        fields[9] = new DBFField();
        fields[9].setName("mes_rojd");
        fields[9].setType(DBFDataType.CHARACTER);
        fields[9].setLength(250);

        // nomer	character 30
        fields[10] = new DBFField();
        fields[10].setName("nomer");
        fields[10].setType(DBFDataType.CHARACTER);
        fields[10].setLength(30);

        // dt_to_gic	character 10
        fields[11] = new DBFField();
        fields[11].setName("dt_to_gic");
        fields[11].setType(DBFDataType.CHARACTER);
        fields[11].setLength(10);

        // mfr	character 2
        fields[12] = new DBFField();
        fields[12].setName("mfr");
        fields[12].setType(DBFDataType.CHARACTER);
        fields[12].setLength(2);

        // y_uch	numeric 20
        fields[13] = new DBFField();
        fields[13].setName("y_uch");
        fields[13].setType(DBFDataType.NUMERIC);
        fields[13].setLength(20);

        // m_uch	numeric 20
        fields[14] = new DBFField();
        fields[14].setName("m_uch");
        fields[14].setType(DBFDataType.NUMERIC);
        fields[14].setLength(20);

        // d_uch	numeric 20
        fields[15] = new DBFField();
        fields[15].setName("d_uch");
        fields[15].setType(DBFDataType.NUMERIC);
        fields[15].setLength(20);

        // svertka	character 7
        fields[16] = new DBFField();
        fields[16].setName("svertka");
        fields[16].setType(DBFDataType.CHARACTER);
        fields[16].setLength(7);

        return fields;
    }


    public static int writeResultSetToDBF(ResultSet rs, String txtDbfFileName)  throws SQLException  {
        // write here code for butch insert to dbf
        int rowsCount = 0;

        ResultSetMetaData md = rs.getMetaData();
        int columns = md.getColumnCount();
        if(columns != 17) throw new SQLException("From MSAccess selected not exact 17 columns. Check mdb file structure.");

        // List<HashMap<String,Object>> list = new ArrayList<HashMap<String,Object>>();
        DBFWriter writer = null;
        try  {
            Charset charset = Charset.forName("CP866");
            File dbfFile = new File(txtDbfFileName);
            if (dbfFile.exists()){
                if(!dbfFile.delete()){
                    throw new SQLException("DBF file exists and can't be removed.");
                }
            }

            writer = new DBFWriter(dbfFile, charset); // opened in sync mode

            if(!dbfFile.exists() || (dbfFile.length() == 0) ) {
                // сделать структуру файла если файла нет или если длина = 0
                DBFField[] fields = prepareFields();
                writer.setFields(fields);
            }

            while (rs.next()) {
                Object[] rowData = new Object[columns];
                for(int i=1; i<=columns; ++i) {
                    rowData[i-1] = rs.getObject(i);
                }
                writer.addRecord(rowData);
                rowsCount++;
            }
            writer.close();
            System.out.println(new StringBuilder().append("To DBF added rows: ").append(rowsCount));

        } catch (DBFException e) {
            e.printStackTrace();
            throw new SQLException("Get DBFException.", e.getMessage());
        }

        return rowsCount;

    }

    /*
    private static List<String> ReadFileLines(String fromFilename, Charset charset)
    {
        List<String> linesList = new ArrayList<String>();
        Path currentRelativePath = Paths.get("");
        String s = currentRelativePath.toAbsolutePath().toString();
        System.out.println("Current absolute path is: " + s);

        try {

            Scanner scanner = new Scanner(new File(fromFilename), charset);
            String nextLine = null;

            while (scanner.hasNextLine()) {
                nextLine = scanner.nextLine();
                if (nextLine.charAt(0) != '#')
                {
                    linesList.add(nextLine);
                }
            }
            scanner.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return linesList;
    }
    */
    private static List<String> ReadFileLinesOld(String fromFilename, Charset charset) {
        List<String> list = new ArrayList<>();
        try (Stream<String> stream = Files.lines(Paths.get(fromFilename), charset)) {
            list = stream
                    .filter(line -> !line.startsWith("#"))
                    .collect(Collectors.toList());
        } catch (IOException e) {
            e.printStackTrace();
        }

        return list;

    }

}
