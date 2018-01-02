package ListMaker;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.net.URL;

/**
 * Created by rahul on 2017-08-19.
 */

public class GUI extends JFrame {
    private JLabel instructionsLabel;
    private JLabel fileOneLabel;

    private JTextField fileOneField;

    private JButton startButton;
    private JButton helpButton;

    private JTextField shortageField;
    private File lastQuantityFile = new File("LastQuantity.txt");

    private JPanel panel;
    private JTextArea area;

    String fileOneString = "";

    File directory;

    GUI() {
        super("TMV Shortage List Maker");
        this.setSize(1000, 300);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setResizable(true);

        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));

        instructionsLabel = new JLabel("Select a File. ");

        panel.add(instructionsLabel);

        JPanel inputPanel = new JPanel();
        inputPanel.setLayout(new BoxLayout(inputPanel, BoxLayout.Y_AXIS));

        String lastQuantityCell = "";
        String lastFileName = "";
        try {
            BufferedReader br = new BufferedReader(new FileReader(lastQuantityFile));
            lastFileName = br.readLine();
            lastQuantityCell = br.readLine();
        } catch (Exception e) {
            //e.printStackTrace();
            lastQuantityCell = "";
            lastFileName = "";
        }
        /*String helpString;
        try{
            URL url = getClass().getResource("/resources/TMV Shortage List Maker READ ME.txt");
            InputStream name = url.openStream();
            BufferedReader txtReader = new BufferedReader(new InputStreamReader(name));

            //BufferedReader txtReader = new BufferedReader(new InputStreamReader(GUI.class.getResourceAsStream("resources/TMV Shortage List Maker READ ME.txt")));

            if (txtReader.readLine().length() > 0){
                helpString = "Help";
            }else{
                helpString = "";
            }

           /* File helpFile = new File("TMV Shortage List Maker READ ME.txt");
            if (helpFile.exists()){
                helpString = "Help";
            }else{
                helpString = "";
            }*/
        // }catch (Exception e){
        //   e.printStackTrace();
        // helpString = "";
        //}

        JPanel fileOnePanel = new JPanel();
        fileOnePanel.setLayout(new FlowLayout());
        fileOneLabel = new JLabel("TMV File (Excel)");
        fileOneField = new JTextField(lastFileName, 40);
        fileOnePanel.add(fileOneLabel);
        fileOnePanel.add(fileOneField);
        JButton fileOneBut = new JButton("Select");
        fileOneBut.addActionListener(new FileOneButListener());
        fileOnePanel.add(fileOneBut);

        JPanel fileOneInputPanel = new JPanel();
        fileOneInputPanel.setLayout(new FlowLayout());



        JLabel fileOneStart = new JLabel("Starting Quantity Cell:");
        shortageField = new JTextField(lastQuantityCell, 5);
        fileOneInputPanel.add(fileOneStart);
        fileOneInputPanel.add(shortageField);



        //adds all pannels
        inputPanel.add(fileOnePanel);
        inputPanel.add(fileOneInputPanel);

        panel.add(inputPanel); //adds to main panel

        JPanel startPanel = new JPanel(new FlowLayout());
        startButton = new JButton("Start"); //to start program
        startButton.addActionListener(new StartButtonListener());
        startPanel.add(startButton);


        helpButton = new JButton("Help");
        helpButton.addActionListener(new HelpButtonListener());
        startPanel.add(helpButton);


        panel.add(startPanel);

        this.add(panel); //adds everything to the gui
        this.setVisible(true);//makes visible
    }

    public class StartButtonListener implements ActionListener {

        String fieldOne; //file name

        //starting cells
        String shortStartCell;

        @Override
        public void actionPerformed(ActionEvent e) {
            //get info from fields
            fieldOne = fileOneField.getText();

            shortStartCell = shortageField.getText().toUpperCase();

            //error check
            if (fieldOne.length() > 4) { //make sure not too short ".xlsx" is 5 characters, so name must be longer than that
                if (fieldOne.substring(fieldOne.length() - 5).equals(".xlsx") || (fieldOne.substring(fieldOne.length() - 4).equalsIgnoreCase(".xls"))) { //makes sure the end of file name has ".xlsx"
                    try {
                        //make file
                        File file = new File(fieldOne); //check if fieldOne and fieldTwo work instead later
                        InputStream fs = new FileInputStream(file); //input stream
                        Workbook wb = WorkbookFactory.create(fs);
                        Sheet sheet = wb.getSheetAt(0);

                        FileOutputStream os = new FileOutputStream("output.txt");

                        PrintWriter output = new PrintWriter(os);
                        String header = "";
                        StringBuilder headBuild = new StringBuilder();
                        headBuild.append("Shortage");
                        for (int i = header.length(); i < 5; i++) {
                            headBuild.append(' ');
                        }
                        headBuild.append("Part Number");

                        for (int k = header.length(); k < 5; k++) {
                            headBuild.append(' ');
                        }

                        headBuild.append("Part Description");

                        header = headBuild.toString();

                        output.println(header);


                        //makes variables
                        Row row;
                        Cell cell;
                        Cell startingShortCell = null;
                        Cell shortCell;


                        int rows; // No of rows
                        rows = sheet.getPhysicalNumberOfRows();


                        int cols = 0;
                        int tmp = 0; //used to get columns


                     //used to count the number of columns
                        for (int i = 0; i < 10 || i < rows; i++) {
                            row = sheet.getRow(i);
                            if (row != null) {
                                tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                                if (tmp > cols) {
                                    cols = tmp;
                                }
                            }
                        }

                        int loopCount = 0; //used to make loop go 1 time
                        for (int r = 0; r < rows && loopCount < 1; r++) { //when it is found, loop will stop
                            row = sheet.getRow(r); //gets row
                            if (row != null) {
                                for (int c = 0; c < cols && loopCount < 1; c++) {
                                    cell = row.getCell(c);
                                    if (cell != null) {
                                        CellRangeAddress range = new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex());
                                        String rangeString = range.toString(); //gets address of cell
                                        if (rangeString.indexOf(shortStartCell) != -1) { //starting shortage cell
                                            startingShortCell = cell;
                                            loopCount++;
                                        }
                                    }
                                }
                            }
                        }

                        if (startingShortCell != null) {
                            for (int r = startingShortCell.getRowIndex(); r < rows; r++) { //starts from row where user entered

                                row = sheet.getRow(r);
                                if (row != null) {
                                    //gets cells
                                    shortCell = row.getCell(startingShortCell.getColumnIndex());

                                    //sets initial values
                                    String shortCellCon = "";

                                    if (shortCell != null) {

                                        //gets values
                                        shortCell.setCellType(Cell.CELL_TYPE_STRING);
                                        if (shortCell.getStringCellValue().length() > 0){

                                            shortCellCon = shortCell.getStringCellValue();
                                            int shortCellInt = Integer.parseInt(shortCellCon);

                                            if (shortCellInt < 0) {
                                                String line = "";
                                                StringBuilder lineBuilder = new StringBuilder();
                                                Cell partNumberCell = row.getCell(0); //will always be this column
                                                partNumberCell.setCellType(Cell.CELL_TYPE_STRING);
                                                Cell partDescriptionCell = row.getCell(1);//will always be this column
                                                partDescriptionCell.setCellType(Cell.CELL_TYPE_STRING);
                                                StringBuilder shortBuilder = new StringBuilder();

                                                shortCellCon = shortCellCon.substring(1);
                                                String spaces = "";
                                                boolean stay = true;

                                                for (int i = 8; i > 0 && stay; i--){
                                                    if (i > shortCellCon.length()){
                                                        shortBuilder.append(' ');
                                                    }else{
                                                        stay = false;
                                                    }
                                                }
                                                shortBuilder.append(shortCellCon);

                                                lineBuilder.append(shortBuilder.toString());

                                                while(spaces.length() < 5){
                                                    lineBuilder.append(' ');
                                                    spaces += 'a';
                                                }
                                                spaces = "";
                                                lineBuilder.append(partNumberCell.getStringCellValue());
                                                while (partNumberCell.getStringCellValue().length() + spaces.length() < 16){
                                                    lineBuilder.append(' ');
                                                    spaces += 'a';
                                                }

                                                lineBuilder.append(partDescriptionCell.getStringCellValue());

                                                line = lineBuilder.toString();
                                                output.println(line);
                                            }
                                        }

                                    }
                                }
                            }
                        }
                        fs.close();
                        output.close();
                        os.close();

                        PrintWriter shortOut = new PrintWriter(lastQuantityFile);
                        shortOut.println(fileOneField.getText());
                        shortOut.println(shortStartCell);
                        shortOut.close();

                        Runtime rt = Runtime.getRuntime();
                        String txtPath = "output.txt";
                        Process p = rt.exec("notepad " + txtPath);

                    } catch (Exception err) {
                        JOptionPane.showMessageDialog(null, "Error");
                        err.printStackTrace();
                    }

                } else {
                    JOptionPane.showMessageDialog(null, "Wrong File Format");
                }

            } else {
                JOptionPane.showMessageDialog(null, "File Name Too Short");
            }


        }
    }

    public class FileOneButListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            FileChooserGUI fcg;
            if (directory != null) {
                fcg = new FileChooserGUI();
            } else {
                fcg = new FileChooserGUI();
            }
            fileOneString = fcg.getPath();
            fileOneField.setText(fileOneString);
            directory = fcg.getDir();
        }
    }

    public class HelpButtonListener implements ActionListener{

        @Override
        public void actionPerformed(ActionEvent e) {
            final String newline = "\n";
            area = new JTextArea(20, 20);
            area.setEditable(false);
            area.append("TMV Shortage List Maker.jar" + newline);
            area.append("--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" + newline);
            area.append("Enter an Excel file. (.xlsx and .xls both work)" + newline);
            area.append("Enter the starting quantity cell to be checked. Ex (Q7) The letter can be lower case or upper case."+newline);
            area.append("The program will make a list which shows Shortages (negative numbers), Part Number, and Part Description in a .txt file called output.txt which will automatically open." + newline);
            area.append("After closing the program and restarting it, the last file name and last starting quantity cell will already be filled in because this information is saved in LastQuantity.txt." + newline);
            panel.add(area);
            panel.revalidate();
            panel.repaint();
            helpButton.setEnabled(false);

        }
    }


}

