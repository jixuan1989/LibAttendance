package fit.lib.thu;

import java.util.HashMap;
import java.util.Map;

/**
 * 
 * @author xd
 *
 */
public class Utils {
	public static String[] upSections = new String[] { "06:00", "12:00",
		"18:30","00:00" };
	public static String[] downSections = new String[] { "11:59", "17:59",
		"23:59", "04:59" };
	public static String[] defaultUpSections = new String[] { "10:15", "15:45",
		"19:59","00:00" };
	public static String[] defaultDpSections = new String[] { "10:16", "15:46",
		"20:00","00:01" };
	public static double[] upTime = new double[4];
	public static double[] downTime = new double[4];
	public static double[] defaultUpTime = new double[4];
	public static double[] defaultDownTime = new double[4];
	static {
		for (int i = 0; i < upSections.length; i++) {
			upTime[i] = timeToDouble(upSections[i]);
		}
		for (int i = 0; i < downSections.length; i++) {
			downTime[i] = timeToDouble(downSections[i]);
		}
		for (int i = 0; i < defaultUpSections.length; i++) {
			defaultUpTime[i] = timeToDouble(defaultUpSections[i]);
		}
		for (int i = 0; i < defaultDpSections.length; i++) {
			defaultDownTime[i] = timeToDouble(defaultDpSections[i]);
		}
	}
	//公用座位集合
	public static Map<String, String> publicUsers=new HashMap<String, String>(20);
	static{
		publicUsers.put("郭雨晨", "刘璋、郭雨晨");
		publicUsers.put("刘璋", "刘璋、郭雨晨");
		publicUsers .put("余志伟", "余志伟、赵凯");
		publicUsers .put("赵凯", "余志伟、赵凯");
		publicUsers.put("宋亮", "宋亮、吕程");
		publicUsers.put("吕程", "宋亮、吕程");
		publicUsers.put("沈晓明", "沈晓明、葛斯涵");
		publicUsers.put("葛斯涵", "沈晓明、葛斯涵");
		publicUsers.put("窦蒙", "窦蒙、王子旋");
		publicUsers.put("王子旋", "窦蒙、王子旋");
		publicUsers.put("董子禾", "董子禾、刘聪");
		publicUsers.put("刘聪", "董子禾、刘聪");
		publicUsers.put("郭芬", "郭芬、徐昊");
		publicUsers.put("徐昊", "郭芬、徐昊");
		publicUsers.put("宋金凤", "宋金凤、龚玉斌");
		publicUsers.put("龚玉斌", "宋金凤、龚玉斌");
		publicUsers.put("殷明", "殷明、钟雨");
		publicUsers.put("钟雨", "殷明、钟雨");
	}
	public static double timeToDouble(String time) {
		if(time==null){
			return -1;
		}
		String[] times = time.split(":");
		return Integer.parseInt(times[0]) + Double.parseDouble(times[1]) / 60;
	}
	
}
