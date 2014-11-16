package org.bitbudget.vegetable.office.poi.utils;

import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {

	private static ExcelUtil instance = null;
	
	private ExcelUtil(){
	}
	
	public static ExcelUtil getInstance(){
		if(null == instance){
			instance = new ExcelUtil();
		}
		return instance;
	}
	
		
	public String getColAliaseByNum(int num){
		String[] str = new String[]{"A","B","C","D","E","F","G","H",
							"I","J","K","L","M","N","O","P","Q","R","S",
							"T","U","V","W","X","Y","Z"};
		if(num>=26){
			return str[num/26-1]+str[num%26];
		}else{
			return str[num];
		}
		
//		String rtn = "";
//		if( num/26 > 0 && num/(26*26) == 0){	//两位
//			rtn =  str[num/26-1]+str[num%26];
//		}else if(num/(26*26) > 0){	//三位
//			
//		}else if(num/26 == 0 && num%26 >= 0){	//一位
//			rtn = str[num];
//		}else{}
//		
//		return rtn;
		
//	      String rtn = "";
////          List<int> iList = new List<int>();
//          List iList = new ArrayList();
//
//          //To single Int
//          while (value / 26 != 0 || value % 26 != 0)
//          {
//              iList.Add(value % 26);
//               value /= 26;
//           }
//
//           //Change 0 To 26
//           for (int j = 0; j < iList.Count - 1; j++)
//           {
//               if (iList[j] == 0)
//               {
//                   iList[j + 1] -= 1;
//                   iList[j] = 26;
//               }
//           }
//
//           //Remove 0 at last
//           if (iList[iList.Count - 1] == 0)
//           {
//               iList.Remove(iList[iList.Count - 1]);
//           }
//
//           //To String
//           for (int j = iList.Count - 1; j >= 0; j--)
//           {
//               char c = (char)(iList[j] + 64);
//               rtn += c.ToString();
//           }
//
//           return rtn;
//       }
	}
}
