import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Scanner;


class BatteryData {
    public String batteryID;
    public Double value;
}

class BatteryComparator implements Comparator<BatteryData> {
    @Override
    public int compare(BatteryData o1, BatteryData o2) {
        return Double.compare(o1.value, o2.value);
    }
}

public class sortBattery {

    static ArrayList<BatteryData> batteries = new ArrayList<>();
    static CSVFromFolder files;

    public static void main(String[] args) throws IOException {
        if(args.length < 2) {
            Scanner scanner = new Scanner(System.in);
            System.out.println("Absolute file path of the folder: ");
            files = new CSVFromFolder(scanner.nextLine());
            System.out.println("Data Field: ");
            getValues(scanner.nextLine());
        } else {
            files = new CSVFromFolder(args[0]);
            getValues(args[1]);
        }
        sortBattery();
        for(BatteryData battery : batteries) {
            System.out.println(battery.batteryID.substring(0, battery.batteryID.length()-4)+"    "+battery.value);
        }
    }

    public static void sortBattery() {
        Collections.sort(batteries, new BatteryComparator());
    }

    public static void getValues(String fieldName) throws IOException {
        while(files.hasNext()) {
            BatteryData battery = new BatteryData();
            battery.batteryID = files.getFileName();
            ArrayList<String> val = files.getData(fieldName);
            Double allValues[] = new Double[val.size()];
            for(int i = 0 ; i < val.size() ; i++) {
                allValues[i] = Double.parseDouble(val.get(i));
            }
            battery.value=doMath(allValues);
            batteries.add(battery);
            files.nextFile();
        }
    }

    public static Double doMath(Double [] values) {
        double number=0;
        for(Double value : values) {
            number+=value;
        }
        return number/values.length;
    }
}
