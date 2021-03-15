using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace CreatEXcelQWWordOT
{
    class DataElemProduc
    {
        delegate string SS();
        SS sS;
        
          public  class ВР114F
        {
            public  string PosOtv = "34.1";
                public  static double Pog  = 0.05;
            public  static string Chert = "14504968.494726.ВР1.1/4К.002.02";
            public static string Ima = "Закладная под ключ ВР 1 1/4F";
        }
            
       public  class ВР2F : ВР114F
        {
            new public static double PosOtv = 50.1;
            
           new  public static string Chert = "14504968.494726.ВР2К.008.02";
           new  public static string Ima = "Закладная под ключ ВР 2F";
            
        }
        public class ZN34 : ВР114F
        {
           new  public static double PosOtv = 15.1;
            public static double BurtNar = 24;
            public static double Hei = 31.4;
            public static double NarDia = 19.9;
       
          new  public static string Chert = "14504968.494726.ZN3/4.001.02";
          new   public static string Ima = "Закладная для накидной гайки ZN3/4";
        }





    }
}
