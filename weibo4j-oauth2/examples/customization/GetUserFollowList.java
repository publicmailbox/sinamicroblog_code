package customization;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.sql.Connection;
import java.sql.DriverManager;
import java.util.ArrayList;
import java.util.List;

import weibo4j.Friendships;
import weibo4j.Timeline;
import weibo4j.Users;
import weibo4j.examples.oauth2.Log;
import weibo4j.model.Status;
import weibo4j.model.StatusWapper;
import weibo4j.model.User;
import weibo4j.model.UserWapper;
import weibo4j.model.WeiboException;
import weibo4j.model.Paging;

import jxl.Workbook;
import jxl.write.*;


public class GetUserFollowList {

	/**
	 * @param args
	 */
	//需要根据个人情况更改的参数
	private static String token = "2.00Tu6dACOvDxqD96d6ca2c9dieqS6C";//获取到的access token
	private static String store_path = "D:/SocialComputing/"; // 存储文件所在的目录
	private static String uid="1841660057";//用户自己的uid
	

	public static User getUser() {
		User user = null;
		Users um=new Users(token);
		try {
			user = um.showUserById(uid);
		} catch (WeiboException e) {
			e.printStackTrace();
		}
		return user;
	}

	public static void getFriendsBilateral() {
		Friendships fm = new Friendships(token);
		try {
			UserWapper users = fm.getFriendsBilateral(uid);
			System.out.println("******************************************");
			System.out.println(uid + " has " + users.getTotalNumber()
					+ " BilateralFriends.");
		} catch (WeiboException e) {
			e.printStackTrace();
		}
	}

	public static List<User> getFollowers() {
		Friendships fm = new Friendships(token);
		List<User> tmp_users = new ArrayList<User>();
		int base_count = 200;
		System.out.println("Getting "+uid+"'s followers...");
		try {
			UserWapper users = fm.getFollowersById(uid, base_count, 0);
			int size = users.getUsers().size();
			int i = 1;
			while (users.getUsers() != null && size > 0) {
				tmp_users.addAll(users.getUsers());
				users =fm.getFollowersById(uid,base_count, i * base_count);
				size = users.getUsers().size();
				i++;
			}
			System.out.println("Followers' count:"+size);
		} catch (WeiboException e) {
			e.printStackTrace();
		}
		return tmp_users;
	}
	
	public static List<Status> getStatus() {
		Timeline tm = new Timeline(token);
		List<Status> tmp_status = new ArrayList<Status>();
		int base_count = 200;
		Paging pg=new Paging();
		pg.setCount(base_count);
		pg.setPage(1);

		try {
			StatusWapper status = tm.getUserTimelineByUid(uid,pg,0,0);
			
			int size = status.getStatuses().size();
			int i = 2;
			while (status.getStatuses() != null && size > 0) {
				System.out.println("Uid" + uid);
				tmp_status.addAll(status.getStatuses());
				pg.setPage(i);
				status = tm.getUserTimelineByUid(uid,pg,0,0);
				size = status.getStatuses().size();
				i++;
			}
		} catch (WeiboException e) {
			e.printStackTrace();
		}
		return tmp_status;
	}
	
	
	public static void InsertToExcel() {
		List <User> followers=getFollowers();
		String username=getUser().getName();
		try   {
			System.out.println("Writing into excel...");
			WritableWorkbook book  =  Workbook.createWorkbook(new File(store_path+"test.xls" ));
	        //  生成名为“第一页”的工作表，参数0表示这是第一页 
	        WritableSheet sheet  = book.createSheet( "sheet1" ,  0 );
	        //  在Label对象的构造子中指名单元格位置是第一列第一行(0,0)
	        // 构造表头，user1和user2 
	        Label label1  =   new  Label( 0 ,  0 ,  " user1 " );
	        Label label2  =   new  Label( 1 ,  0 ,  " user2 " );
	        //  将定义好的单元格添加到工作表中 
	        sheet.addCell(label1);
	        sheet.addCell(label2);
	        label1=null;
	        label2=null;
	        User follower=null;
	        // 将user的followers写入excel中，user1为当前user，user2为对应followers的名称
	        int follower_count=followers.size();
	        int i=1;
	        while(i<=follower_count){
	        	label1=new  Label( 0,i,username);
	        	follower=followers.get(i-1);
	        	label2=new  Label( 1,i,follower.getName());
		        sheet.addCell(label1);
		        sheet.addCell(label2);
		        label1=null;
		        label2=null;
		        follower=null;
		        i++;
	        }
	        System.out.println("Rows count:"+i);
	        //  写入数据并关闭文件 
	        book.write();
	        book.close();
	        } catch  (Exception e)  {
	        	System.out.println(e);
	        } 

	}
	

	public static void main(String[] args) {

		// TODO Auto-generated method stub
		InsertToExcel();
	}

}
