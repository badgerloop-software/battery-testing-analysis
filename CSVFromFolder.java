import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Objects;

public class CSVFromFolder {

    ArrayList<File> files;
    int currentFile;
    BufferedReader reader;
    String [] fieldNames;

    public CSVFromFolder(String folderPath) throws IOException {
        File folder = new File(folderPath);
        files = new ArrayList<File>(Arrays.asList(Objects.requireNonNull(folder.listFiles())));
        filterFile();
        currentFile=0;
        String name = files.get(currentFile).getPath();
        reader = new BufferedReader(new FileReader(files.get(currentFile).getPath()));
        initFieldNames();
    }

    public void filterFile() {
        for (int i = 0; i < files.size(); i++) {
            if(!files.get(i).getName().endsWith(".csv")) {
                files.remove(i);
                i--;
            }
        }
    }

    public ArrayList<String> getData(String fieldName) throws IOException {

        //find the field name index
        int index = -1;
        for (int i = 0; i < fieldNames.length; i++) {
            if(fieldNames[i].equals(fieldName)) {
                index = i;
                break;
            }
        }
        if(index==-1)
            throw new IOException("Field name not found");

        String line;
        ArrayList<String> data = new ArrayList<>();
        while ((line = reader.readLine()) != null){
            data.add(line.split(",")[index]);
        }
        return data;
    }

    public String[] getFieldNames() {
        return fieldNames;
    }

    public String getFileName() {
        String name = files.get(currentFile).getName();
        return name;
    }

    private void initFieldNames() throws IOException {
        //reader.reset();
        String line = reader.readLine();
        fieldNames = line.split(",");
    }

    public boolean hasNext() {
        return currentFile+1<files.size();
    }

    public void nextFile() throws IOException {
        if(currentFile+1<files.size()) {
            currentFile++;
            //System.out.println(currentFile+" "+files.size());
            reader = new BufferedReader(new FileReader(files.get(currentFile).getPath()));
            initFieldNames();
        }
    }
}
