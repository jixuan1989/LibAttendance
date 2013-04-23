package fit.lib.thu;

import java.text.Collator;
import java.text.RuleBasedCollator;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

public class User  implements Comparator<User>{

	String type;
	String name;
	List<String[][]> records;//签到和签退时间 [4][2] -表示无记录，up和down应同时为-  ,三个时段，签到签退  最后是加班
	double[][] time=null;//统计计算后的时间
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public List<String[][]> getRecords() {
		return records;
	}
	public void setRecords(List<String[][]> records) {
		this.records = records;
	}

	public double[][] getTime() {
		if(time==null)
			calculateTime();
		return time;
	}
	public void setTime(double[][] time) {
		this.time = time;
	}
	public User(String type, String name, int days) {
		super();
		this.type = type;
		this.name = name;
		this.records = new ArrayList<String[][]>();
		for(int i=0;i<days;i++){
			records.add(new String[4][2]);
		}
	}
	/**
	 * just for generate comparator
	 */
	public User(){}
	public void addTime(String time,int dayIndex) {
		double t = Utils.timeToDouble(time);
		String[][] times=records.get(dayIndex);
		for (int i = 0; i < 4; i++) {
			if (t >= Utils.upTime[i] && t <= Utils.defaultUpTime[i]) {
				times[i][0] = time;
			} else if (t >= Utils.defaultDownTime[i] && t <= Utils.downTime[i]) {
				times[i][1] = time;
			}
		}
	}
	public void fillTime(int day){
		String[][] times=records.get(day);
		for(int i=0;i<4;i++){
			if(times[i][0]==null&&times[i][1]==null){
				continue;
			}
			if(times[i][0]==null){
				times[i][0]=Utils.defaultUpSections[i];
			}
			if(times[i][1]==null){
				times[i][1]=Utils.defaultDpSections[i];
			}
		}
		if(times[3][0]!=null){
			times[2][1]=Utils.downSections[2];
		}
	}

	@Override
	public String toString() {
		String str= "User [type=" + type + ", name=" + name + ", records=[\t" ;
		for(String[][] re:records){
			str+="上午签到:"+re[0][0]+"\t"+"上午签退:"+re[0][1]+";\t";
			str+="下午签到:"+re[1][0]+"\t"+"下午签退:"+re[1][1]+";\t";
			str+="晚上签到:"+re[2][0]+"\t"+"晚上签退:"+re[2][1]+";\t";
			str+="加班签到:"+re[3][0]+"\t"+"加班签退:"+re[3][1]+";\t";
		}
		return str+"]]";
	}
	private static Map<String, Integer> levels=new HashMap<String, Integer>(10);
	static{
		levels.put("博士生", 1);
		levels.put("研究生", 2);
		levels.put("本科生", 3);
		levels.put("外协",  4);
		levels.put("公用座位统计",  5);
	}
	public int getLevel(){
		Integer integer= levels.get(type);
		return integer==null?0:integer;
	}
	public String toString(int day,int flag,int upDown){
		if(records.get(day)[flag][upDown]==null){
			return "-";
		}
		else return records.get(day)[flag][upDown];
	}
	/**
	 * 
	 * @param day
	 * @param flag
	 * @return
	 */
	public double getLong(int day,int flag){
		double up=Utils.timeToDouble(records.get(day)[flag][0]);
		double down=Utils.timeToDouble(records.get(day)[flag][1]);
		if(up==-1||down==-1){
			return 0;
		}
		double result=down-up;
		switch (flag) {
		case 3:
			result*=0.85;
			break;
		default:
			break;
		}
		return result;
	}
	public int getOnlineDays(){
		int result=0;
		boolean online=false;
		for(String[][] day:records){
			for(int i=0;i<day.length;i++){
				online=online||day[i][0]!=null;
			}
			if(online)
				result++;
		}
		return result;
	}
	static RuleBasedCollator collator = (RuleBasedCollator)Collator.getInstance(Locale.CHINA);
	@Override
	public int compare(User o1, User o2) {
		int result= o1.getLevel()-o2.getLevel();
		if(result==0){
			result=(int) (o2.calculateTime()-o1.calculateTime());
		}
		if(result==0){
			result= collator.compare(o1.getName(), o2.getName());
		}
		return result;
	}
	/**
	 * 返回最后的总时间。同时计算出所有需要的时间
	 * @return
	 */
	public double calculateTime(){
		if(time==null){
			time=new double[records.size()+1][5];
			for(int day=0;day<records.size();day++){
				for(int j=0;j<4;j++){
					time[day][j]=getLong(day, j);
					time[day][4]+=time[day][j];
					time[records.size()][j]+=time[day][j];
				}
				time[records.size()][4]+=time[day][4];
			}
		}
		return time[records.size()][4];

	}


}
