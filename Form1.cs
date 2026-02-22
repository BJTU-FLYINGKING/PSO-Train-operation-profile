using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MathNet.Numerics;
using NPOI;
using MathNet.Numerics.Integration;

using static liziqun0117.Form1;
using DocumentFormat.OpenXml.Bibliography;
using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.Atp;
using static ICSharpCode.SharpZipLib.Zip.ExtendedUnixData;
using System.IO;

namespace liziqun0117
{
    public partial class Form1 : Form
    {
        List<int> preceding_position = new List<int> { 0, 10, 25, 46, 84, 210, 284, 600, 1200, 1245, 1800, 2325, 2400, 3000, 3105, 3600, 4200, 4593, 4800, 5400, 6000, 6600, 6745, 7200, 7800, 7945, 8400, 9000, 9135, 9600, 9745, 10200, 10395, 10800, 10995, 11395, 11400, 12000, 12185, 12600, 12795, 13200, 13395, 13800, 14400, 15000, 15600, 16200, 16800, 17400, 18000, 18600, 19045, 19200, 19645, 19800, 20095, 20400, 20495, 21000, 21600, 21645, 22200, 22245, 22800, 23400, 24000, 24600, 25200, 25800, 26400, 27000, 27600, 28200, 28800, 29400, 30000, 30600, 31200, 31800, 32400, 33000, 33600, 34200, 34800, 35400, 36000, 36600, 37200, 37800, 38400, 39000, 39600, 39865, 40200, 40465, 40795, 40800, 41225, 41400, 42000, 42465, 42600, 43092, 43200, 43695, 43800, 44395, 44400, 45000, 45225, 45600, 45855, 46200, 46655, 46800, 47400, 47993, 48000, 48600, 49200, 49800, 50400, 51000, 51600, 52200, 52800, 53400, 54000, 54600, 55200, 55800, 56335, 56400, 57000, 57355, 57577, 57600, 57615, 57637, 57675, 57700, 57800, 57900, 58000, 58100, 58200, 58300, 58375 };
        List<double> preceding_speed = new List<double>{0, 10.95453265723628, 13.849241757869299, 18.59528234857822, 27.73692346456411, 48.25345555544717, 49.849057021926825, 62.84665513784734, 71.86925144177106, 73.67130418339063, 85.78254931914461, 122.98737992262174, 126.46955179814806, 153.6892647611202, 157.4956333913965, 169.99551578383543, 168.99083942456284, 166.96325153417783, 167.2292894844642, 167.9546062871437, 170.66126569900575, 166.98344211435796, 168.72929675960106, 169.7644581466371, 168.98386056871314, 168.8040991230396, 166.9263450282767, 168.7910380779079, 167.74023345383694, 168.95235176729398, 168.7996395974586, 170.99129551228228, 171.887778503683, 171.82480566958156, 170.4452571422433, 167.60898551904197, 167.57347283077922, 167.67254415265964, 167.70217322072608, 169.45263523817724, 170.26054342518515, 176.31511201948717, 179.12886406101728, 178.92629739995294, 189.22855278174592, 198.80805827978392, 198.85473819227064, 198.9020253320684, 198.94992665619824, 198.9984491760294, 199.0475999569969, 199.09738611829465, 199.13472597038475, 199.83762749933976, 198.45476616145507, 198.86445996719416, 197.37414149766946, 198.19281169571903, 198.4454279892044, 199.76959014201122, 196.73989756652156, 196.86277211600702, 198.69313833645504, 198.14789181285582, 198.18301799763543, 198.22147784009522, 198.260448294936, 198.29993549516584, 198.33994562980533, 198.3804849439051, 198.42155973854702, 198.46317637082808, 198.50534125382686, 198.5480608565518, 198.59134170387094, 198.6351903764223, 198.67961351050494, 198.72461779794946, 198.77020998596785, 198.81639687698208, 198.86318532843057, 198.91058225255242, 198.95859461614853, 199.00722944031912, 199.05649380017684, 199.10639482453536, 199.15693969557256, 199.20813564846753, 199.25998997101104, 199.31251000318903, 199.36570313673795, 199.41957681467204, 199.47413853078132, 199.498458210496, 198.45531032733976, 199.66028552236975, 197.99777993752085, 198.01114929662202, 199.13615238659634, 196.8994343609642, 198.5119850856981, 199.73097603382172, 199.31334999025572, 197.77326890168382, 197.36547189906045, 199.3156712276896, 198.71945839654725, 195.34673036992868, 195.21925702449195, 194.17522915639032, 193.78764514956626, 195.82606844694152, 197.185822076875, 198.10274548708264, 199.3029091639804, 199.31567298270824, 199.3689066162036, 199.42218831968225, 199.3470431058735, 197.46576390268, 195.5401425682475, 198.3048804582996, 196.39927467157395, 199.14940480633828, 197.26355715176933, 199.99898043345863, 198.13262324692715, 196.22293997697957, 198.97606925322592, 197.08619767408436, 199.82463988572349, 197.9543143826479, 182.8353491669152, 182.4825467606278, 139.42188137140033, 99.31900009305012, 49.48424179312221, 49.7369064141893, 49.08176454552234, 49.32583393240928, 49.7440421358884, 47.23850757667334, 48.386379026321215, 49.50278342482983, 44.94165227559015, 46.156863240389555, 47.33568645319885, 48.480840148730756, 0.0,0};
        List<double> preceding_travelTime_raw = new List<double>{1259.6281790998814, 1252.4981790998813, 1248.1440030703563, 1243.4837405402627, 1237.5785608666804, 1225.64020746327, 1220.2091539050873, 1200.0202757623738, 1167.9527930895179, 1165.7266094762406, 1140.666067430849, 1122.5600115743919, 1120.3953092497259, 1104.9754839768486, 1102.5460601696898, 1091.663323311073, 1078.9194475366508, 1070.4968681540781, 1066.0371635913818, 1053.1487162211547, 1040.390897236885, 1027.5963830625715, 1024.4865814997777, 1014.8084118067874, 1002.0555809842375, 998.9648846759594, 989.2070564626697, 976.3390900796671, 973.450799801097, 963.5070108762499, 960.4159854294966, 950.7747629556187, 946.6800248791427, 938.1961923566474, 934.0941683921037, 925.5748255276935, 925.46742133901, 912.5813622028812, 908.6096853560636, 899.7472876531868, 895.6143894078959, 887.2006420653174, 883.2506524903583, 875.1066559393745, 863.3724631184245, 852.2394929891834, 841.376017606711, 830.5151086388241, 819.656798208496, 808.801118772776, 797.9481031244338, 787.0977843935973, 779.0522252768407, 776.255038979859, 768.2106975588458, 765.4018730001304, 760.0414664366414, 754.489941006441, 752.7654476733137, 743.6347021352287, 732.739628665573, 731.9164635171726, 721.8142253900033, 720.9977775547629, 710.9152936982488, 700.0173345545948, 689.1215037102651, 678.2278285561496, 667.336336788481, 656.4470564109783, 645.560015736934, 634.6752433913219, 623.7927683128221, 612.9126197558784, 602.0348272927082, 591.1594208152636, 580.2864305371813, 569.4158869957168, 558.5478210536172, 547.6822639009569, 536.8192470569846, 525.9588023718808, 515.100962028485, 504.2457585440093, 493.3932247716985, 482.5433939024266, 471.6962994662664, 460.85197533401384, 450.01045571865006, 439.1717751767643, 428.33596860991986, 417.5030712659498, 406.67311874022914, 401.8908354065284, 395.82982985590326, 391.03725197176465, 385.0622694267679, 384.9713623856827, 377.26641264277663, 374.0848803710236, 363.15955098613796, 354.7526227366384, 352.3168031130777, 343.3958276093325, 341.4279112303306, 332.44336517425245, 330.54403533189986, 319.67276516936715, 319.58059124444304, 308.4864433987201, 304.31078611991984, 297.38084524315656, 292.7092308613298, 286.4252140598621, 278.18174806570937, 275.56270307694365, 264.72706954342334, 254.020712119153, 253.8943232304721, 243.0075778339031, 232.01537662554904, 221.04659513174607, 210.10168881684146, 199.18015064939763, 188.28242425149915, 177.4080034559737, 166.55732001345046, 155.6027389123663, 144.67153741868142, 133.76416158226317, 122.88010503556073, 112.0198019610913, 101.90398115835927, 100.62290512781946, 87.20277286971904, 76.49660480867584, 65.75490335285203, 64.08590432399582, 62.99299345046231, 61.383361649158665, 58.6216745364042, 56.76567053843254, 49.236249844955324, 41.88099230703025, 34.25746218945024, 26.35393054656955, 18.652782141057855, 11.138420834774607, 0};
        List<double> preceding_travelTime = new List<double>();
        public static int index_preceding_arrSwitch;
        public static double preceding_arrSwitchTime;


        public static double t_start = 49.42;
        public static int switchPosition_arr2_start = 57637;
        public static int switch_time=10;
        public static int switch_length = 38;
        public static int train_length = 200;
        public struct FenXiangQv
        {
            public int start_pos;
            public int end_pos;
            public int length;
            public FenXiangQv(int s,int e,int l)
            {
                start_pos = s;
                end_pos = e;
                length = l;
            }
        }
        public struct Grid
        {
            public int grid;
            public int start_pos;
            public int end_pos;
            public int length;
            public Grid(int g,int s, int e, int l)
            {
                grid= g;    
                start_pos = s;
                end_pos = e;
                length = l;
            }
        }
        public Grid[] GRID;
        public FenXiangQv[] FXQ;
        public static int grid_qd_num = 30;
        public static int fxq_qd_num = 9;
        public static LinkList_SLOPE<int> slope_ori = new LinkList_SLOPE<int>();
        public static LinkList_SLOPE<int> slope_final_qd = new LinkList_SLOPE<int>();
        public static LinkList_FENXIANGQV<int> fxq_ori = new LinkList_FENXIANGQV<int>();
        public static Best_Set<int> best_cz = new Best_Set<int> ();
        public static Best<int> best;
        public static Best<int> delete;
        //存储的运行日志
        //全线线路的表格
        public static string filepath = "E:\\曲线规划\\川藏工程数据\\全线下行线数据表.xlsx";
        //对应区段线路坡度的sheet表
        public static string sheetname = "邦达-夏里";
        //方向,0为下行，1为上行
        public static int fx = 0;        
        //StreamWriter writer = new StreamWriter(txt_date_path);
        Best<int> bestA;
        public Form1()
        {
            InitializeComponent();
            //前车数据
            double total_traveltime_pre = 21;
            foreach (int travel_time in preceding_travelTime_raw)
            {
                preceding_travelTime.Add(total_traveltime_pre * 60 -0.37 - travel_time-t_start);
            }

            index_preceding_arrSwitch = preceding_position.FindLastIndex(pos => pos == switchPosition_arr2_start);
            preceding_arrSwitchTime = preceding_travelTime[index_preceding_arrSwitch];

            int[,] qd_BD_XL = new int[grid_qd_num,2];
            FXQ = new FenXiangQv[fxq_qd_num];
            GRID = new Grid[grid_qd_num];
            
            //int fx = 0; //0下行1上行
            Read_Excel(filepath,sheetname,fx);
            slope_ori.QD_LAST(fx);
            SLOPE<int> s = slope_ori.Head;
            double v1 = 10;
            int m = 400;
            double a1 = Acc(10, 0, v1, m);
            Console.WriteLine(a1);

            int[] gk = new int[112];
            double[] speed = new double[112];
            //double[] travelTime = new double[100];
            double t=0;
            Console.WriteLine("over:" + 0 + "time:" + DateTime.Now);
            int MODE = 0;
            int T_ = 1260;
            int T_error = 10;
            int M = 400;
            double e = 0;     //能耗
            double s_e = 0;  //停车误差
            (gk, speed, t, e, s_e) =GK_LZQ(slope_ori, fxq_ori, 200/3.6, T_, T_error,M,MODE);
            Console.WriteLine("over:" + 1 + "time:" + DateTime.Now);
            string write_txt = "result:"+t.ToString("0.00")+"s\ne:"+e.ToString("0.00")+
                "kJ\ns:"+s_e.ToString("0.00")+"\n";
            for(int i = 0;i<slope_ori.QD_NUM();i++)
            {
                Console.WriteLine(gk[i].ToString() + "gk and speed" + speed[i].ToString()+" "+i.ToString());
                write_txt += gk[i].ToString() + " " + speed[i].ToString() + " " + slope_ori.FindSlope(i).ToString()+"\n";
            }
            write_txt += speed[slope_ori.QD_NUM()];
            Console.WriteLine("vs_over:"+"time:"+DateTime.Now);
        }
        public void console_shuzu()
        {
            for(int i=0;i<grid_qd_num;i++)
            {
                Console.WriteLine(GRID[i].start_pos+" " + GRID[i].end_pos+" "+i);
            }
            for (int i = 0; i < fxq_qd_num; i++)
            {
                Console.WriteLine(FXQ[i].start_pos + " " + FXQ[i].end_pos+" " + FXQ[i].length);
            }
        }
        public int RunTimeMin(int[] qd, int xs_speed)
        {
            int T = 0;
            int length = 0;
            for (int i = 1; i <= qd[0]; i++)
            {
                length += qd[i];
            }
            T =(int) (length / xs_speed);
            return (T+xs_speed);
        }
        public static float[] l=new float[5] { 0,0,0,0,0};
        Best_Set<int> bestset=new Best_Set<int>();
        public int sudu(float[,] speed, int[] qd,int m)
        {
            int i = qd[0];
            int k = 0;
            int re = 0;
            int[] flag= new int[qd[0]-2];
            for(int j = 1;j<i-1;j++)
            {
                if (Math.Abs(speed[m,j] - speed[m,j+1])<=3.5)
                {
                    flag[j-1] = 1;
                }
                else
                {
                    flag[j - 1] = 0;
                }
            }
            for(int j = 0; j < qd[0]-2;j++)
            {
                if (flag[j]==0)
                {
                    k = 1;
                    break;
                }
            }
            if(k==0)
            {
                re = 1;
            }
            if(k==1)
            {
                re = 0;
            }
            return re;
        }
        public float Wz(float sz, float v, float a, float c,float qd)
        {
            double E=0;
            float e = 0;
            float v0 = v*3.6f;
            float v1 = 2f * a * sz + v * v;
            if(v1<0)
            {
                sz = sz - 1;
            }
            float v2 =(float)Math.Sqrt(v1)*3.6f;
            if (c == 0)
            {
                E = Integrate.OnClosedInterval(s => (0.00807f * (2*a*s+v*v)*3.6f*3.6f + 0.0622f *3.6f* (float)Math.Sqrt(v*v+2*a*s) + 2.031f), 0, sz);
            }
            else if(c==1)
            {
                E = (2.031f + 0.0622f * v0 + 0.00807f * v0 * v0) * qd;
            }
            Console.WriteLine(v.ToString()+" "+E.ToString("0.00")+" "+sz.ToString()+" "+v0.ToString("0.00")+" "+a.ToString()+" "+c.ToString()+" "+qd.ToString("0.00"));
            return (float)E;    
        }
        public double Acc(int gk,int grid,double v_start,int Mass)
        {
            double a = 0;
            double a_max = 0;
            double a_fb = 0;
            double v_c=v_start*3.6;
            if(v_c > 50)
            {
                a_fb = -(-0.043 * v_start*3.6 + 586.2)/Mass;
            }
            else if ((v_c <= 50) && (v_c >= 0))
            {
                a_fb = -(0.0035*v_start * 3.6 * v_start * 3.6 - 2.5805*v_start * 3.6 + 704.325) / Mass;
            }
            int flag = 0;
            if ((v_c * 3.6) > 160)
            {
                a_max = (31512.0 / (v_start * 3.6)) / Mass;
                //Console.WriteLine("228");
            }
            else if((v_c <= 160)&&(v_c >= 0))
            {
                a_max = 197.0/Mass;
                //Console.WriteLine("232");
            }
            double a_w = 0;
            a_w = 9.8 * (0.16 + 0.0053 * v_start * 3.6 + 0.000118 * v_start * 3.6 * 3.6*v_start)/1000;
            double g = 9.8;
            double a_grid = g * grid * 0.001;
            //Console.WriteLine("299:" + a_max + " " + a_w + " " + a_grid + " " + v_start * 3.6);
            switch (gk)
            {
                case 1:
                    a = a_max*0.1;
                    break;
                case 2:
                    a = a_max * 0.2;
                    break;
                case 3:
                    a = a_max * 0.3;
                    break;
                case 4:
                    a = a_max * 0.4;
                    break;
                case 5:
                    a = a_max * 0.5;
                    break;
                case 6:
                    a = a_max * 0.6;
                    break;
                case 7:
                    a = a_max * 0.7;
                    break;
                case 8:
                    a = a_max * 0.8;
                    break;
                case 9:
                    a = a_max * 0.9;
                    break;
                case 10:
                    a = a_max;
                    break;
                case 0:
                    break;
                case -8:
                    a = -1.2;
                    break;
                case -7:
                    a = a_fb;
                    break;
                case -6:
                    a = a_fb*6/7;
                    break;
                case -5:
                    a = a_fb*5/7;
                    break;
                case -4:
                    a = a_fb*4/7;
                    break;
                case -3:
                    a = a_fb*3/7;
                    break;
                case -2:
                    a = a_fb*2/7;
                    break;
                case -1:
                    a = a_fb*1/7;
                    break;
                default:
                    break;
            }
            return a-a_grid-a_w;
        }
        public (int,double,double) Change_V(int gk,double v_start,int dis,double speed_limit,int grid,int Mass,int qd_num,int all_num)
        {
            double t_change = 0;
            double t_yunsu = 0;
            double change_dis = 0;
            int keda = 1;
            double a = Acc(gk,grid,v_start,Mass);
            int gk_linshi = gk;
            //(a,keda) = Acc1(v_start, g,1);
            double v_end = 0;
            v_end = Math.Sqrt(v_start * v_start + 2 * a * dis);

            //Console.WriteLine("299:"+a+" sd:");
            if (a==0)
            {
                t_yunsu=dis/v_start;
                t_change = 0;
            }
            else
            {
                if (qd_num < all_num)
                {
                    if ((v_start * v_start + 2 * a * dis) > 0)
                    {
                        v_end = Math.Sqrt(v_start * v_start + 2 * a * dis);
                        while (v_end > speed_limit)
                        {
                            gk_linshi--;
                            a = Acc(gk_linshi, grid, v_start, Mass);
                            v_end = Math.Sqrt(v_start * v_start + 2 * a * dis);
                            //v_end =speed_limit;
                            //t_yunsu=(dis-(v_end*v_end-v_start*v_start)/(2*a))/v_end;
                        }
                        t_change = (v_end - v_start) / a;
                    }
                    else
                    {
                        v_end = 0;
                        t_yunsu = (dis - (v_end * v_end - v_start * v_start) / (2 * a)) / v_start;
                        t_change = (v_end - v_start) / a;
                    }
                }
                if(qd_num==all_num)
                {
                    v_end = 0;
                    t_change= (v_end - v_start) / a;
                }
            }
            //Console.WriteLine(keda.ToString());
            return (gk_linshi,v_end, t_change+t_yunsu);
        }
        public double Energy(double v_start, double v_end, int slope, int length, double weight,int gk)
        {
            double E_v = 0;
            double E_h = 0;
            double E = 0;
            double E_basic = 0;
            double m = weight;
            double g = 9.8;
            if (gk > 0)
            {
                //E_v = 0.5 * m * (v_end * v_end - v_start * v_start);
                //E_h = m * length * slope * 0.001 * g;
                //E_basic=  9.8 * (0.16 + 0.0053 * v_start * 3.6 + 0.000118 * v_start * 3.6 * 3.6 * v_start) / 1000*m*length;
                E = Acc(gk,slope,v_start,(int)weight)*m*length; //= E_v + E_h+E_basic;
            }
            else
            {
                // E_h = m * length * slope * 0.001 * g;
                // E_basic = 9.8 * (0.16 + 0.0053 * v_start * 3.6 + 0.000118 * v_start * 3.6 * 3.6 * v_start) / 1000 * m * length;
                E = 0;// E_v + E_h + E_basic;
            }
            return E;
        }
        public int Choose(double t1, double e01, double t2, double e02,double T, double dT,int mode)//1表明优于，2表明不劣于，-1表明劣于
        {
            //0为劣于，1为优于，2为不劣于
            double e1 =e01 * 138 * 9.8;
            double e2 =e02 * 138 * 9.8;
            int flag = 0;
            if (mode == 0)   //普通式
            {
                if (((t1<t2)&&(t2 - t1 < dT/2) && (e1 - e2 < 1000) && (e1 > e2) && (Math.Abs(t1 - T) < dT)) || ((t1>t2)&&(t1 - t2 < dT) && (e2 - e1 < 1000) && (e2 > e1) && (Math.Abs(t2 - T) < dT)))
                {
                    flag = 2;
                }
                else if (((t1 < t2) && (Math.Abs(t1 - T) < dT) && (e1 <= e2))|| ((t1 <= t2) && (Math.Abs(t1 - T) < dT) && (e1 < e2)))
                {
                    flag = 1;
                }
                else 
                {
                    flag = 0;
                }
            }
            if(mode==1)  //节时模式
            {
                if ((Math.Abs(t2 - T) > dT)&& (Math.Abs(t1 - T) < dT))
                {
                    flag = 1;
                }
                else if((Math.Abs(t2 - T) < dT) && (Math.Abs(t1 - T) < dT)&&(t1<t2)&&(t1<T))
                {
                    flag = 1;
                }
            }
            if(mode==2)
            {
                if((e1<e2)&&(Math.Abs(t1-T)<dT))
                {
                    flag = 1;
                }
            }
            return flag;
        }
        public bool Include_FXQ(SLOPE<int> qd_slope,LinkList_FENXIANGQV<int> fxq_list)
        {
            FENXIANGQV<int> qd_fxq=fxq_list.Head;
            while (qd_fxq != null)
            {
                ///下行线分相区
                if ((qd_fxq.Start_pos >= qd_slope.Start_pos) && (qd_fxq.Start_pos <= qd_slope.End_Pos))
                {
                    return true;
                }
                if ((qd_fxq.End_Pos >= qd_slope.Start_pos) && (qd_fxq.End_Pos <= qd_slope.End_Pos))
                {
                    return true;
                }
                ///上行线分相区
                if ((qd_fxq.Start_pos <= qd_slope.Start_pos) && (qd_fxq.Start_pos >= qd_slope.End_Pos))
                {
                    return true;
                }
                if ((qd_fxq.End_Pos <= qd_slope.Start_pos) && (qd_fxq.End_Pos >= qd_slope.End_Pos))
                {
                    return true;
                }
                qd_fxq = qd_fxq.Next;
            }
            return false;
        }
        public (int[], double[],double,double,double) GK_LZQ(LinkList_SLOPE<int> slope, LinkList_FENXIANGQV<int> fxq,
            double speed_limit, int T_target, int dT, int WEIGHT, int mode)
        {
            int change_dangwei = 4;
            int qd_num = slope.QD_NUM();
            int xsize = 30000; //种群
            int seed = 0; //随机数种子
            int[] gk = new int[qd_num];                 //工况序列
            int[] best_gk = new int[qd_num];                 //工况序列
            double[] travelTime = new double[qd_num + 1];
            double[] acc = new double[qd_num + 1];   //加速度
            int[,] gk_all = new int[xsize, qd_num];      //工况的粒子群
            double[,] travelTime_all = new double[xsize, qd_num+1];  //运行时间 
            double[,] acc_all = new double[xsize, qd_num + 1];
            //double[,] travelTime_all = new double[xsize, qd_num + 1];  //运行时间 
            Best_Set<int> gk_best = new Best_Set<int>(); //最优工况筛选
            float[,] all_a = new float[xsize, qd_num + 1];  //自身加速度
            float[,] a_best = new float[xsize, qd_num + 1]; //加速度最优
            float[] best_a = new float[qd_num];        //群体加速度最优
            int[,] gk_self_best = new int[xsize, qd_num + 1]; //自身最优
            double[,] v = new double[xsize, qd_num + 1];      //变坡点速度 
            double[] v_send = new double[qd_num + 1];        //变坡点速度
            int[] min_gk_need_break = new int[qd_num];        //变坡点制动限制速度
            int[] max_gk_need_break = new int[qd_num];        //变坡点制动限制速度
            double[] max_speed_need_break = new double[qd_num + 1];        //变坡点制动限制速度
            double[] min_speed_need_break = new double[qd_num + 1];        //变坡点制动限制速度
            double[] t_all = new double[xsize];
            double[] t_pre_all = new double[xsize];
            double[] E_all = new double[xsize];
            double[] t_last = new double[xsize];
            double[] E_last = new double[xsize];
            double[] S_error = new double[xsize];
            double[,] v_all = new double[xsize, qd_num + 1];  //粒子群各粒子的向量速度
            double t_linshi = 0;
            bool NotQvjianTingche = true;
            int MaxIt = 10;                        //迭代次数
            double c1 = 0.8;                           //算法参数
            double c2 = 0.8;                           //算法参数
            double wmax = 1.2;                         //惯性因子
            double wmin = 0.1;                         //惯性因子
            ///min代表最大制动档位 max代表最低制动档位
            double min_break_dis = (-speed_limit * speed_limit) / (2 * Acc(-5, 0,speed_limit,WEIGHT));
            double max_break_dis = (-speed_limit * speed_limit) / (2 * Acc(-4, 0,speed_limit,WEIGHT));
            double min_break_pos = slope.Head.Start_pos + slope.DISTANCE() - min_break_dis;
            double max_break_pos = slope.Head.Start_pos + slope.DISTANCE() - max_break_dis;
            int min_break_gk = slope.BREAK_POS(min_break_pos);
            int max_break_gk = slope.BREAK_POS(max_break_pos);
            double[] v_show = new double[qd_num + 1];
            int[] gk_show = new int[qd_num];
            int num_fxq = 0;
            int pos_fxq = 0;
            for(int i=0;i<qd_num; i++)
            {
                if(Include_FXQ(slope.Find(i),fxq))
                {
                    num_fxq++;
                }
                if(num_fxq==2)
                {
                    pos_fxq = i;
                    break;
                }
            }
            //Console.WriteLine(min_break_gk+ " " + max_break_gk+" "+min_break_dis+" "+max_break_dis);
            for (int i = 0; i < qd_num; i++)
            {
                if (i < min_break_gk)
                {
                    max_gk_need_break[i] = 20;
                }
                else
                {
                    max_gk_need_break[i] = -5;
                }
                if (i < max_break_gk)
                {
                    min_gk_need_break[i] = 20;
                }
                else
                {
                    min_gk_need_break[i] = -1;
                }
            }
            max_speed_need_break[qd_num] = 0;
            min_speed_need_break[qd_num] = 0;
            double[] a = new double[qd_num];
            bool flag_gk_break = false;
            for(int i=0;i<qd_num; i++)
            {
                max_gk_need_break[i] = 20;
                max_speed_need_break[i] = speed_limit;
            }
            for (int i = qd_num - 1; i >= 0; i--)
            {
                if (i >= qd_num - 8)
                {
                    speed_limit = 50 / 3.6;
                }
                else
                {
                    speed_limit = 200 / 3.6;
                }
                    a[i] = Acc(-5, slope.FindSlope(i), max_speed_need_break[qd_num],WEIGHT);
                max_speed_need_break[i] = Math.Sqrt(max_speed_need_break[i + 1] * max_speed_need_break[i + 1] -
                    2 * a[i] * slope.FindLength(i));
                max_gk_need_break[i] = -5;
                Console.WriteLine(max_speed_need_break[i].ToString()+" 544 "+a[i].ToString());
                if ((max_speed_need_break[i] > speed_limit) && (max_speed_need_break[i+1]>=speed_limit))
                {
                    max_speed_need_break[i] = speed_limit;
                    max_gk_need_break[i] = 20;
                }
                else if((max_speed_need_break[i] > speed_limit) && (max_speed_need_break[i+1] < speed_limit))
                {
                    while (max_speed_need_break[i]>speed_limit)
                    {
                        //Console.WriteLine("299:" + a[i]);
                        max_gk_need_break[i]++;
                        a[i] = Acc(max_gk_need_break[i], slope.FindSlope(i), speed_limit, WEIGHT);
                        max_speed_need_break[i] = Math.Sqrt(max_speed_need_break[i + 1] * max_speed_need_break[i + 1] -
                    2 * a[i] * slope.FindLength(i));
                        if (max_speed_need_break[i] <= speed_limit)
                        {
                            max_gk_need_break[i]--;
                            a[i] = Acc(max_gk_need_break[i], slope.FindSlope(i), speed_limit, WEIGHT);
                            max_speed_need_break[i] = speed_limit;
                            Console.WriteLine(i + " 546");
                            Console.WriteLine(max_speed_need_break[i].ToString() + " 565 " + a[i].ToString());
                            if (max_speed_need_break[i] > 50 / 3.6)
                            {       
                                flag_gk_break = true;
                            }
                            break;
                        }
                    }
                }
                if (flag_gk_break)
                {
                    break;
                }
                Console.WriteLine(max_speed_need_break[i] + " " + max_gk_need_break[i]+" "+i+" "+slope.FindLength(i)+" "+a);
            }
            /*int xs_start = 0;
            for (int i = pos_fxq - 1; i > pos_fxq - 1 - 6; i--)
            {
                max_speed_need_break[i] = 120 / 3.6;
                xs_start = i;
            }
            Console.WriteLine(xs_start+" " + max_speed_need_break[xs_start]);
            while ((max_speed_need_break[xs_start]<40))
            {
                xs_start--;
                max_speed_need_break[xs_start] = Math.Sqrt(max_speed_need_break[xs_start+1]*
                    max_speed_need_break[xs_start + 1]-2*Acc(-7, slope.FindSlope(xs_start),max_speed_need_break[xs_start + 1], WEIGHT)*slope.FindSlope(xs_start));
                
            }*/
            max_speed_need_break[0] = 0;
            int gk_max = 4;
            flag_gk_break = false;
            for (int i=0;i<qd_num;i++)
            {
                if (i <= 3)
                {
                    speed_limit = 50 / 3.6;
                    gk_max = 4;
                    max_gk_need_break[i] = gk_max;
                }
                else
                {
                    speed_limit = 200 / 3.6;
                    gk_max += 4; //gk_max * (i + 1);
                    max_gk_need_break[i] = gk_max;
                }
                if(gk_max>10)
                {
                    gk_max = 10;
                    max_gk_need_break[i] = 20;
                }
                a[i] = Acc(gk_max, slope.FindSlope(i), max_speed_need_break[i], WEIGHT);
                max_speed_need_break[i+1]= Math.Sqrt(2 * a[i] * slope.FindLength(i)+
                    max_speed_need_break[i] * max_speed_need_break[i]);
                if(i+1>2)
                {
                    speed_limit = 200 / 3.6;
                }
                while (max_speed_need_break[i+1]>speed_limit)
                {
                    //max_gk_need_break[i]--;
                    gk_max--;
                    a[i] = Acc(max_gk_need_break[i], slope.FindSlope(i), max_speed_need_break[i], WEIGHT);
                    max_speed_need_break[i + 1] = Math.Sqrt(2 * a[i] * slope.FindLength(i) +
                        max_speed_need_break[i] * max_speed_need_break[i]);
                    if (max_speed_need_break[i] <= speed_limit)
                    {
                        //max_gk_need_break[i]--;
                        //a[i] = Acc(max_gk_need_break[i], slope.FindSlope(i), speed_limit, WEIGHT);
                        //max_speed_need_break[i] = speed_limit;
                        Console.WriteLine(i + " 546");
                        Console.WriteLine(max_speed_need_break[i].ToString() + " 565 " + a[i].ToString());
                        if (max_speed_need_break[i] > 50 / 3.6)
                        {
                            flag_gk_break = true;
                        }
                        //break;
                    }
                }
                if (flag_gk_break)
                {
                    break;
                }
                //if (max_speed_need_break[i+1]>speed_limit)
                //{
                //    max_speed_need_break[i + 1] = speed_limit;
                //    gk_max -= 4;
                //    //break;
                //    if (max_speed_need_break[i+1]>50/3.6)
                //    {
                //        break;
                //    }    
                //}
            }
            for (int i = 0; i < qd_num; i++)
            {
                if (max_speed_need_break[i] == speed_limit)
                { 
                    min_speed_need_break[i] = speed_limit/2; 
                }
                else
                {
                    min_speed_need_break[i] = max_speed_need_break[i] / 3;
                    if (min_speed_need_break[i]>speed_limit/2)
                    {
                        min_speed_need_break[i] = speed_limit / 2;
                    }
                }
            }
            for (int i=0;i<qd_num;i++)
            {
                Console.WriteLine(max_speed_need_break[i] + " " + max_gk_need_break[i] + " " + i + " " + slope.FindLength(i) + " " + a[i]+" "+
                    min_speed_need_break[i]);
            }
            string write_length = "";
            for(int i=0;i<slope.QD_NUM();i++)
            {
                Console.WriteLine(slope.FindLength(i));
                write_length += slope.FindLength(i) + "\n";
            }
            string txtpath = "E:\\中期\\粒子群\\单车\\length1.txt";
            File.WriteAllText(txtpath, write_length);
            Console.WriteLine(slope.QD_NUM()+" "+slope.DISTANCE()+" "+
                min_break_gk+" "+min_break_dis);
            double dis_error = 0.5;
            //Console.WriteLine();
            ///初始化种群
            best = best_cz.Head;
            int sc_num = 0;
            while (best == null)
            {   
                sc_num++;
                for (int i = 0; i < xsize; i++)
                {
                    v[i, 0] = 0;
                    v[i, qd_num] = 0;
                    for (int j = 0; j < qd_num; j++)
                    {
                        byte[] bytes = new byte[4];
                        //创建加密服务，实现加密随机数生成器
                        System.Security.Cryptography.RNGCryptoServiceProvider rng = new System.Security.Cryptography.RNGCryptoServiceProvider();
                        //加密数据存入字节数组
                        rng.GetBytes(bytes);
                        //转成整型数据返回，作为随机数生成种子
                        seed = BitConverter.ToInt32(bytes, 0);
                        Random ran = new Random(seed);
                        //seed += 1;
                        if (j == 0)
                        {
                            gk_all[i, j] = ran.Next(1, 5);
                        }
                        else if (j == qd_num - 1)
                        {
                            gk_all[i, j] = ran.Next(-7, 0);
                        }
                        else if(slope.FindSlope(j)>6) ///注意
                        {
                            gk_all[i, j] = ran.Next(0, 11);
                        }
                        else if (slope.FindSlope(j) >= 0) ///注意
                        {
                            gk_all[i, j] = ran.Next(0, 11);
                        }
                        else
                        {
                            gk_all[i, j] = ran.Next(-7, 11);
                        }
                        if (Include_FXQ(slope.Find(j), fxq))
                        {
                            gk_all[i, j] = 0;
                        }
                        /*else if (gk_all[i, j] > min_gk_need_break[j])
                        {
                            gk_all[i,j] = -8;
                        }*/
                        v[i, j] = 0;
                        v_all[i, j] = 0;
                    }
                    for (int j = 0; j < qd_num - 1; j++)
                    {
                        if (Include_FXQ(slope.Find(j), fxq))
                        {
                            gk_all[i, j] = 0;
                            int k = j;
                            while (gk_all[i, k - 1] > gk_all[i, k] + change_dangwei)
                            {
                                gk_all[i, k - 1] = gk_all[i, k] + change_dangwei;
                                k--;
                            }
                            while (gk_all[i, k - 1] < gk_all[i, k] - change_dangwei)
                            {
                                gk_all[i, k - 1] = gk_all[i, k] - change_dangwei;
                                k--;
                            }
                        }
                        if (gk_all[i, j] - change_dangwei > gk_all[i, j + 1])
                        {
                            gk_all[i, j + 1] = gk_all[i, j] - change_dangwei;
                        }
                        if (gk_all[i, j] + change_dangwei < gk_all[i, j + 1])
                        {
                            gk_all[i, j + 1] = gk_all[i, j] + change_dangwei;
                        }
                    }
                    E_all[i] = 0;
                    t_all[i] = 0;
                    t_pre_all[i] = 0;
                }
                for (int i = 0; i < xsize; i++)
                {
                    for (int j = 0; j < qd_num; j++)
                    {
                        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j),WEIGHT,j,qd_num);
                        
                        if (Include_FXQ(slope.Find(j), fxq))
                        {
                            gk_all[i, j] = 0;
                            (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                        }
                        if ((!Include_FXQ(slope.Find(j), fxq)) && (j != 0))
                        {
                            while (v[i, j + 1] > max_speed_need_break[j + 1])
                            {
                                if (gk_all[i, j] == -7)
                                {
                                    break;
                                }
                                //Console.WriteLine(DateTime.Now + " 816 " + gk_all[i, j] + " " + i + " " + j);
                                gk_all[i, j]--;
                                if (j > 0)
                                {
                                    if (gk_all[i, j] > gk_all[i, j - 1] + change_dangwei)
                                    {
                                        gk_all[i, j] = gk_all[i, j - 1] + change_dangwei;
                                        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                                        if(v[i, j + 1] > max_speed_need_break[j + 1])
                                        {
                                            j--;
                                            gk_all[i, j]++;
                                            t_all[i] = travelTime_all[i,j];
                                            break;
                                        }
                                        break;
                                    }
                                    if (gk_all[i, j] < gk_all[i, j - 1] - change_dangwei)
                                    {
                                        gk_all[i, j] = gk_all[i, j - 1] - change_dangwei;
                                        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT,j,qd_num);
                                        if (v[i, j + 1] > max_speed_need_break[j + 1])
                                        {
                                            j--;
                                            gk_all[i, j]--;
                                            t_all[i] = travelTime_all[i, j];
                                            break;
                                        }
                                        break;
                                    }
                                }
                                //gk_all[i, j] = max_gk_need_break[j]; 
                                (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);

                            }
                            if (max_speed_need_break[j + 1] != min_speed_need_break[j + 1])
                            {
                                while (v[i, j + 1] < min_speed_need_break[j + 1])
                                {
                                    if (gk_all[i, j] == 10)
                                    {
                                        break;
                                    }
                                    //Console.WriteLine(DateTime.Now + " 816 " + gk_all[i, j] + " " + i + " " + j);

                                    gk_all[i, j]++;
                                    /*if (j > 0)
                                    {
                                        if (gk_all[i, j] > gk_all[i, j - 1] + change_dangwei)
                                        {
                                            gk_all[i, j] = gk_all[i, j - 1] + change_dangwei;
                                            (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT);
                                            break;
                                        }
                                        if (gk_all[i, j] < gk_all[i, j - 1] - change_dangwei)
                                        {
                                            gk_all[i, j] = gk_all[i, j - 1] - change_dangwei;
                                            (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT);
                                            break;
                                        }
                                    }*/
                                    //gk_all[i, j] = max_gk_need_break[j]; 
                                    (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);

                                }
                            }
                            if (Include_FXQ(slope.Find(j), fxq))
                            {
                                gk_all[i, j] =0;
                                (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                            }
                        }
                        /*(gk_all[i,j],v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j));
                        if (v[i, j + 1] > max_speed_need_break[j + 1])
                        {
                            gk_all[i, j] = max_gk_need_break[j];
                            (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j));
                        }*/
                    }
                }
                ///初始化运行时间和运行能耗
                for (int i = 0; i < xsize; i++)
                {
                    double t_pre_linshi = 0;
                    bool come_to_next_lizi = false;
                    for (int j = 0; j < qd_num; j++)
                    {
                        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT,j,qd_num);

                        if (Include_FXQ(slope.Find(j), fxq))
                        {
                            gk_all[i, j] = 0;
                            (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                        }
                        if ((!Include_FXQ(slope.Find(j), fxq))&&(j!=0))
                        {
                            while (v[i, j + 1] > max_speed_need_break[j + 1])
                            {
                                if (gk_all[i, j] == -7)
                                {
                                    t_all[i] = 0;
                                    E_all[i] = 0;
                                    t_pre_all[i] = 0;
                                    come_to_next_lizi = true;
                                    break;
                                }
                                //Console.WriteLine(DateTime.Now + " 816 " + gk_all[i, j] + " " + i + " " + j);
                                gk_all[i, j]--;
                                if (j > 0)
                                {
                                    if (gk_all[i, j] > gk_all[i, j - 1] + change_dangwei)
                                    {
                                        gk_all[i, j] = gk_all[i, j - 1] + change_dangwei;
                                        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT,j,qd_num);
                                        if (v[i, j + 1] > max_speed_need_break[j + 1])
                                        {
                                            j--;
                                            gk_all[i, j]++;
                                            t_all[i] = travelTime_all[i, j]; //-= t_pre_linshi;
                                            break;
                                        }
                                        break;
                                    }
                                    if (gk_all[i, j] < gk_all[i, j - 1] - change_dangwei)
                                    {
                                        gk_all[i, j] = gk_all[i, j - 1] - change_dangwei;
                                        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                                        if (v[i, j + 1] > max_speed_need_break[j + 1])
                                        {
                                            j--;
                                            gk_all[i, j]--;
                                            t_all[i] = travelTime_all[i, j];
                                            break;
                                        }
                                        break;
                                    }
                                }
                                //gk_all[i, j] = max_gk_need_break[j]; 
                                (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);

                            }
                            if (max_speed_need_break[j + 1] != min_speed_need_break[j + 1])
                            {
                                while (v[i, j + 1] < min_speed_need_break[j + 1])
                                {
                                    if (gk_all[i, j] == 10)
                                    {
                                        t_all[i] = 0;
                                        E_all[i] = 0;
                                        t_pre_all[i] = 0;
                                        come_to_next_lizi = true;
                                        break;
                                    }
                                    //Console.WriteLine(DateTime.Now + " 816 " + gk_all[i, j] + " " + i + " " + j);

                                    gk_all[i, j]++;
                                    /*if (j > 0)
                                    {
                                        if (gk_all[i, j] > gk_all[i, j - 1] + change_dangwei)
                                        {
                                            gk_all[i, j] = gk_all[i, j - 1] + change_dangwei;
                                            (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT);
                                            break;
                                        }
                                        if (gk_all[i, j] < gk_all[i, j - 1] - change_dangwei)
                                        {
                                            gk_all[i, j] = gk_all[i, j - 1] - change_dangwei;
                                            (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT);
                                            break;
                                        }
                                    }*/
                                    //gk_all[i, j] = max_gk_need_break[j]; 
                                    (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);

                                }
                            }
                            if (Include_FXQ(slope.Find(j), fxq))
                            {
                                gk_all[i, j] = 0;
                                (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                            }
                        }
                        if(j==qd_num-1)
                        {
                            double a_stop = Acc(gk_all[i, j],slope.FindSlope(j), v[i, j], WEIGHT);
                            S_error[i] = slope.FindLength(j) + v[i,j]*v[i,j]/(2*a_stop);
                            double s1 = S_error[i];
                            while (Math.Round(S_error[i])>dis_error)
                            {
                                if (S_error[i] < 0)
                                {
                                    gk_all[i, j]--;
                                }
                                if (S_error[i] > 0)
                                {
                                    gk_all[i, j]++;
                                }
                                if (S_error[i] == 0)
                                {
                                    break;
                                }
                                if ((gk_all[i, j] >= -1) || ((s1 < 0) && (S_error[i]>0)))
                                {
                                    Console.WriteLine(s1+" 975 " + S_error[i]+" " + gk_all[i,j]+" "+a_stop);
                                    break;
                                } 
                                a_stop = Acc(gk_all[i, j], slope.FindSlope(j), v[i, j], WEIGHT);
                                s1= S_error[i];
                                S_error[i] = slope.FindLength(j) + v[i, j] * v[i, j] / (2 * a_stop);
                            }
                        }
                        if(true == come_to_next_lizi)
                        {
                            break;
                        }
                        //Console.WriteLine("v:" + v[i,j]+" " + gk_all[i,j]+" "
                        //   +Acc(gk_all[i, j], slope.FindSlope(j),v[i,j],WEIGHT)
                        //   +" "+t_linshi+" "+ slope.FindSlope(j));
                        t_pre_all[i] = t_all[i];
                        travelTime_all[i, j] = t_all[i];
                        t_all[i] += t_linshi;
                        t_pre_linshi = t_linshi;
                        // travelTime_all[i, j] = t_all[i];
                        acc_all[i, j] = Acc(gk_all[i, j], slope.FindSlope(j), v[i, j], WEIGHT);

                        ////撞软墙约束
                        //int index_preceding = preceding_travelTime.FindLastIndex(time => time < t_all[i]);
                        ////前车在该时刻下的状态
                        //int preceding_currentPosition = preceding_position[index_preceding];
                        //double preceding_currentSpeed = (preceding_speed[index_preceding] / 3.6 + preceding_speed[index_preceding + 1] / 3.6) / 2;
                        //double preceding_currentTime = preceding_travelTime[index_preceding];
                        //double preceding_Braking = -1.2;   //待填充

                        //// 列车进站约束
                        //if (t_all[i] >= 1198 && t_pre_all[i] < 1198)
                        //{
                        //    preceding_currentPosition = switchPosition_arr2_start;
                        //    preceding_currentSpeed = (preceding_speed[index_preceding_arrSwitch] / 3.6 + preceding_speed[index_preceding_arrSwitch + 1] / 3.6) / 2;
                        //    preceding_currentTime = preceding_arrSwitchTime;

                        //    //计算上个时刻后车的状态
                        //    double v_tracking = v[i, j - 1];
                        //    double delta_trac_t = preceding_arrSwitchTime - (t_pre_all[i] + t_start);
                        //    double trac_a = acc_all[i, j - 1];
                        //    double position_tracking = slope.Find(j - 1).Start_pos - slope.Head.Start_pos + v_tracking * delta_trac_t + 0.5 * trac_a * delta_trac_t * delta_trac_t;
                        //    double t_A = (switch_length + train_length) / preceding_currentSpeed;
                        //    double L_safe = (t_A + switch_time) * v_tracking;
                        //    while (preceding_currentPosition - position_tracking < L_safe)
                        //    {
                        //        if (gk_all[i, j] == -7)
                        //        {
                        //            break;
                        //        }
                        //        gk_all[i, j]--;
                        //        trac_a = Acc(gk_all[i, j], slope.FindSlope(j), v[i, j], WEIGHT);
                        //        position_tracking = slope.Find(j - 1).Start_pos - slope.Head.Start_pos + v_tracking * delta_trac_t + 0.5 * trac_a * delta_trac_t * delta_trac_t;
                        //        t_A = (switch_length + train_length) / preceding_currentSpeed;
                        //        L_safe = (t_A + switch_time) * v_tracking;
                        //    }
                        //}

                        E_all[i] += Energy(v[i, j], v[i, j + 1], slope.FindSlope(j), slope.FindLength(j), WEIGHT, gk_all[i, j]);
                    }
                }
                /*for(int i=0;i<xsize;i++)
            {
                for(int j=0;j<qd_num;j++)
                {
                    Console.Write(gk_all[i, j]+" ");
                }
                Console.WriteLine();
            }
            Console.WriteLine(" speed:");
            for (int i = 0; i < xsize; i++)
            {
                for (int j = 0; j < qd_num+1; j++)
                {
                    Console.Write(v[i, j] + " ");
                }
                Console.WriteLine();
            }*/
                double t_best = 0;
                for (int i = 0; i < xsize; i++)
                {
                    if(mode==1)
                    {
                        if ((Math.Abs(S_error[i]) <= dis_error) && (Math.Abs(t_all[i] - T_target) < dT) && (t_all[i]<T_target))
                        {
                            for (int j = 0; j < qd_num; j++)
                            {
                                gk[j] = gk_all[i, j];
                                v_send[j] = v[i, j];
                                travelTime[j] = travelTime_all[i, j];
                                acc[j] = acc_all[i, j];
                                //Console.Write(gk[j] + " ");
                            }
                            v_send[qd_num] = v[i, qd_num];
                            travelTime[qd_num] = travelTime_all[i, qd_num];
                            best_cz.Add_Best(gk, v_send, t_all[i], E_all[i], S_error[i], travelTime,acc);
                        }
                    }
                    else if ((Math.Abs(S_error[i])<=dis_error)&&(Math.Abs(t_all[i] - T_target) < dT))
                    {
                        for (int j = 0; j < qd_num; j++)
                        {
                            gk[j] = gk_all[i, j];
                            v_send[j] = v[i, j];
                            travelTime[j] = travelTime_all[i, j];
                            acc[j] = acc_all[i, j];
                            //Console.Write(gk[j] + " ");
                        }
                        t_best = t_all[i];
                        v_send[qd_num] = v[i, qd_num];
                        travelTime[qd_num] = travelTime_all[i, qd_num];
                        best_cz.Add_Best(gk, v_send, t_all[i], E_all[i], S_error[i], travelTime,acc);
                    }
                    //Console.WriteLine(t_all[i]+" " + S_error[i]);
                //Console.WriteLine("647"+t_all[i]+" "+i);
                }
                best=best_cz.Head;
                Console.WriteLine(sc_num.ToString()+" "+T_target.ToString()+" "+t_best.ToString());

            }
            best = best_cz.Head; 
           
            int flag_ori = 0;
            ///筛选初始非劣解集
          
            ///种群迭代
            for (int iter = 0; iter < MaxIt; iter++)
            {
                double w = wmax - (wmax - wmin) * iter / MaxIt;
                Console.WriteLine("680:"+iter.ToString());
                //从非劣解中选择粒子作为全局最优解
                for (int i = 0; i < xsize; i++)
                {
                    //E_all[i] = 0;
                    for (int j = 1; j < qd_num; j++)
                    {
                        Best<int> best_ran = best_cz.Return_Best();
                        byte[] bytes = new byte[4];
                        //创建加密服务，实现加密随机数生成器
                        System.Security.Cryptography.RNGCryptoServiceProvider rng = new System.Security.Cryptography.RNGCryptoServiceProvider();
                        //加密数据存入字节数组
                        rng.GetBytes(bytes);
                        //转成整型数据返回，作为随机数生成种子
                        seed = BitConverter.ToInt32(bytes, 0);
                        Random ran1 = new Random(seed);
                        seed += 1;
                        Random ran2 = new Random(seed);
                        seed += 1;
                        Random ran3 = new Random(seed);
                        seed += 1;
                        Random ran4 = new Random(seed);
                        best_gk[j] = best_ran.GK[j];
                        double x1 = c1 * ran1.Next(0, 1) * (gk_self_best[i, j] - gk_all[i, j]);
                        double x2 = c2 * ran2.Next(0, 1) * (best_gk[j] - gk_all[i, j]);
                        //bestA = bestset.Head;//bestset.Return_Best();
                        v_all[i, j] = w * v_all[i, j] + c1 * (ran1.Next(0, 10)) * 0.1f * (gk_self_best[i, j] - gk_all[i, j]) + c2 * (ran2.Next(0, 10) * 0.1f) * (best_ran.GK[j] - gk_all[i, j]);
                        gk_all[i, j] = gk_all[i, j] + (int)(Math.Round(v_all[i, j]));
                        if (gk_all[i,j]>10)
                        {
                            gk_all[i, j] = 10;
                        }
                        else if (gk_all[i, j] < -7)
                        {
                            gk_all[i, j] = -7;
                        }
                        if (Include_FXQ(slope.Find(j), fxq))
                        {
                            gk_all[i, j] = 0;
                        }
                        if (gk_all[i,0] >change_dangwei)
                        {
                            
                            gk_all[i, j] = change_dangwei;
                        }
                        if (gk_all[i, 0] < 0)
                        {
                            gk_all[i, j] = ran1.Next(1,change_dangwei);
                        }
                        /*else if (gk_all[i, j] > max_gk_need_break[j])
                        {
                            gk_all[i, j] = -8;
                        }*/
                    }
                    for (int j = 0; j < qd_num - 1; j++)
                    {
                        if (Include_FXQ(slope.Find(j), fxq))
                        {
                            gk_all[i, j] = 0;
                            int k = j;
                            while (gk_all[i, k-1] > gk_all[i,k]+change_dangwei)
                            {
                                gk_all[i, k-1] = gk_all[i, k] + change_dangwei;
                                k--;
                            }
                            while (gk_all[i, k - 1] < gk_all[i, k] - change_dangwei)
                            {
                                gk_all[i, k - 1] = gk_all[i, k] - change_dangwei;
                                k--;
                            }
                            //(gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT);
                        }
                        if (gk_all[i, j] - change_dangwei > gk_all[i, j + 1])
                        {
                            gk_all[i, j + 1] = gk_all[i, j] - change_dangwei;
                        }
                        if (gk_all[i, j] + change_dangwei < gk_all[i, j + 1])
                        {
                            gk_all[i, j + 1] = gk_all[i, j] + change_dangwei;
                        }
                    }
                }
                //Console.WriteLine("line547 "+DateTime.Now+" "+iter.ToString());
                ///计算时间
                for (int i = 0; i < xsize; i++)
                {
                    t_all[i] = 0;
                    E_all[i] = 0;
                    bool come_to_next_lizi = false;
                    for (int j = 0; j < qd_num; j++)
                    {
                        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                        if ((!Include_FXQ(slope.Find(j), fxq)) && (j != 0))
                        {
                            while (v[i, j + 1] > max_speed_need_break[j + 1])
                            {
                                if (gk_all[i, j] == -7)
                                {
                                    t_all[i] = 0;
                                    E_all[i] = 0;
                                    t_pre_all[i] = 0;
                                    come_to_next_lizi = true;
                                    break;
                                }
                                //Console.WriteLine(DateTime.Now + " 816 " + gk_all[i, j] + " " + i + " " + j);
                                gk_all[i, j]--;
                                //if (j > 0)
                                //{
                                //    if (gk_all[i, j] > gk_all[i, j - 1] + change_dangwei)
                                //    {
                                //        gk_all[i, j] = gk_all[i, j - 1] + change_dangwei;
                                //        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                                //        if (v[i, j + 1] > max_speed_need_break[j + 1])
                                //        {
                                //            j--;
                                //            gk_all[i, j]++;
                                //            t_all[i]=travelTime_all[i,j];
                                //            break;
                                //        }
                                //        break;
                                //    }
                                //    if (gk_all[i, j] < gk_all[i, j - 1] - change_dangwei)
                                //    {
                                //        gk_all[i, j] = gk_all[i, j - 1] - change_dangwei;
                                //        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                                //        if (v[i, j + 1] > max_speed_need_break[j + 1])
                                //        {
                                //            j--;
                                //            gk_all[i, j]--;
                                //            t_all[i] = travelTime_all[i, j];
                                //            break;
                                //        }
                                //        break;
                                //    }
                                //}
                                //gk_all[i, j] = max_gk_need_break[j]; 
                                (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);

                            }
                            if (max_speed_need_break[j + 1] != min_speed_need_break[j + 1])
                            {
                                while (v[i, j + 1] < min_speed_need_break[j + 1])
                                {
                                    if (gk_all[i, j] == 10)
                                    {
                                        break;
                                    }
                                    //Console.WriteLine(DateTime.Now + " 816 " + gk_all[i, j] + " " + i + " " + j);
                                    gk_all[i, j]++;
                                    //if (gk_all[i, j] > gk_all[i, j - 1] + change_dangwei)
                                    //{
                                    //    gk_all[i, j] = gk_all[i, j - 1] + change_dangwei;
                                    //    (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                                    //    break;
                                    //}
                                    //if (gk_all[i, j] < gk_all[i, j - 1] - change_dangwei)
                                    //{
                                    //    gk_all[i, j] = gk_all[i, j - 1] - change_dangwei;
                                    //    (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                                    //    break;
                                    //}
                                    //if (j > 0)
                                    //{
                                    //    if (gk_all[i, j] > gk_all[i, j - 1] + change_dangwei)
                                    //    {
                                    //        gk_all[i, j] = gk_all[i, j - 1] + change_dangwei;
                                    //        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                                    //        break;
                                    //    }
                                    //    if (gk_all[i, j] < gk_all[i, j - 1] - change_dangwei)
                                    //    {
                                    //        gk_all[i, j] = gk_all[i, j - 1] - change_dangwei;
                                    //        (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                                    //        break;
                                    //    }
                                    //}
                                    //gk_all[i, j] = max_gk_need_break[j]; 
                                    (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);

                                }
                            }
                            while ((v[i, j + 1] == 0) && (j != qd_num - 1))
                            {
                                //Console.WriteLine(DateTime.Now+" 733 " + gk_all[i,j] + " "+ i+" "+j);
                                if (gk_all[i, j] == 10)
                                {
                                    break;
                                }
                                if (gk_all[i, j] > gk_all[i, j - 1] + change_dangwei)
                                {
                                    gk_all[i, j] = gk_all[i, j - 1] + change_dangwei;
                                    (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                                    break;
                                }
                                if (gk_all[i, j] < gk_all[i, j - 1] - change_dangwei)
                                {
                                    gk_all[i, j] = gk_all[i, j - 1] - change_dangwei;
                                    (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT,j,qd_num);
                                    break;
                                }
                                gk_all[i, j]++;
                                //Console.WriteLine("545 " + DateTime.Now+" " + gk_all[i,j]+" "+ v[i, j]+" "+ v[i, j + 1]+" "+i+" "+j);
                                (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT,j,qd_num);
                            }
                            if (Include_FXQ(slope.Find(j), fxq))
                            {
                                gk_all[i, j] = 0;
                                (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j), WEIGHT, j, qd_num);
                            }
                        }
                        if (j == qd_num - 1)
                        {
                            double a_stop = Acc(gk_all[i, j], slope.FindSlope(j), v[i, j], WEIGHT);
                            S_error[i] = slope.FindLength(j) + v[i, j] * v[i, j] / (2 * a_stop);
                            double s1 = S_error[i];
                            while (Math.Abs(S_error[i]) > 0.5)
                            {
                                if (S_error[i] < 0)
                                {
                                    gk_all[i, j]--;
                                }
                                if (S_error[i] > 0)
                                {
                                    gk_all[i, j]++;
                                }
                                if (S_error[i] == 0)
                                {
                                    break;
                                }
                                if ((gk_all[i, j] > -1) || ((s1 < 0) && (S_error[i] > 0)))
                                {
                                    break;
                                }
                                a_stop = Acc(gk_all[i, j], slope.FindSlope(j), v[i, j], WEIGHT);
                                s1 = S_error[i];
                                S_error[i] = slope.FindLength(j) + v[i, j] * v[i, j] / (2 * a_stop);
                            }
                        }
                        if(come_to_next_lizi == true)
                        {
                            break;
                        }
                        /*while ((v[i, j + 1] != 0) && (j == qd_num - 1))
                        {
                            Console.WriteLine(DateTime.Now + " 744");
                            //Console.WriteLine(v[i, j + 1] + " " + j + " " + gk_all[i, j]);
                            if (gk_all[i, j] == -8)
                            {
                                break;
                            }
                            gk_all[i, j]--;
                            (gk_all[i, j], v[i, j + 1], t_linshi) = Change_V(gk_all[i, j], v[i, j], slope.FindLength(j), speed_limit, slope.FindSlope(j));
                        }*/
                        t_pre_all[i] = t_all[i];
                        travelTime_all[i, j] = t_all[i];
                        t_all[i] += t_linshi;
                        E_all[i] += Energy(v[i, j], v[i, j + 1], slope.FindSlope(j), slope.FindLength(j), WEIGHT, gk_all[i, j]);
                        acc_all[i,j]= Acc(gk_all[i, j], slope.FindSlope(j), v[i, j], WEIGHT);


                        ////撞软墙约束
                        //int index_preceding = preceding_travelTime.FindLastIndex(time => time < t_all[i]);
                        ////前车在该时刻下的状态
                        //int preceding_currentPosition = preceding_position[index_preceding];
                        //double preceding_currentSpeed = (preceding_speed[index_preceding] / 3.6 + preceding_speed[index_preceding + 1] / 3.6) / 2;
                        //double preceding_currentTime = preceding_travelTime[index_preceding];
                        //double preceding_Braking = Acc(-7, slope.FindSlope(j), v[i, j], WEIGHT);   //待填充

                        //// 列车进站约束
                        //if (t_all[i] >= 1198 && t_pre_all[i] < 1198)
                        //{
                        //    preceding_currentPosition = switchPosition_arr2_start;
                        //    preceding_currentSpeed = (preceding_speed[index_preceding_arrSwitch] / 3.6 + preceding_speed[index_preceding_arrSwitch + 1] / 3.6) / 2;
                        //    preceding_currentTime = preceding_arrSwitchTime;

                        //    //计算上个时刻后车的状态
                        //    double v_tracking = v[i, j - 1];
                        //    double delta_trac_t = preceding_arrSwitchTime - (t_pre_all[i] + t_start);
                        //    double trac_a = acc_all[i, j - 1];
                        //    double position_tracking = slope.Find(j - 1).Start_pos - slope.Head.Start_pos + v_tracking * delta_trac_t + 0.5 * trac_a * delta_trac_t * delta_trac_t;
                        //    double t_A = (switch_length + train_length) / preceding_currentSpeed;
                        //    double L_safe = (t_A + switch_time) * v_tracking;
                        //    while (preceding_currentPosition - position_tracking < L_safe)
                        //    {
                        //        if (gk_all[i, j] == -7)
                        //        {
                        //            break;
                        //        }
                        //        gk_all[i, j]--;
                        //        trac_a = Acc(gk_all[i, j], slope.FindSlope(j), v[i, j], WEIGHT);
                        //        position_tracking = slope.Find(j - 1).Start_pos - slope.Head.Start_pos + v_tracking * delta_trac_t + 0.5 * trac_a * delta_trac_t * delta_trac_t;
                        //        t_A = (switch_length + train_length) / preceding_currentSpeed;
                        //        L_safe = (t_A + switch_time) * v_tracking;
                        //    }
                        //}
                    }
                }
                for (int i = 0; i < xsize; i++)
                {
                    if (Math.Abs(S_error[i]) <= dis_error)
                    //if ((t_all[i] < t_last[i]) && (Math.Abs(t_last[i] - T_target) < dT) && (t_all[i] > 0))
                    {
                        int flag_self = 0;
                        flag_self = Choose(t_all[i], E_all[i], t_last[i], E_last[i], T_target, dT, mode);
                        if (flag_self == 1)
                        {
                            t_last[i] = t_all[i];
                            E_last[i] = E_all[i];
                            for (int j = 0; j < qd_num; j++)
                            {
                                gk_self_best[i, j] = gk_all[i, j];
                            }
                        }
                        best = best_cz.Head;
                        //flag_ori = Choose(t_all[i], E_all[i], best.T_zu, best.E, T_target, dT, mode);
                        int add_flag = 0;
                        for (int j = 0; j < qd_num; j++)
                        {
                            gk[j] = gk_all[i, j];
                            v_send[j] = v[i, j];
                            travelTime[j] = travelTime_all[i, j];
                            acc[j] = acc_all[i, j];
                        }
                        v_send[qd_num] = v[i, qd_num];
                        travelTime[qd_num] = travelTime_all[i, qd_num];
                        while (best != null)
                        {
                            //Console.WriteLine(DateTime.Now + " 782 "+iter);
                            flag_ori = Choose(t_all[i], E_all[i], best.T_zu, best.E, T_target, dT, mode);
                            if (flag_ori == 0)
                            {
                                //Console.WriteLine("786");
                                best = best.Next;
                            }
                            else if (flag_ori == 1)
                            {
                                Console.WriteLine("791");
                                delete = best;
                                best = best.Next;
                                if (add_flag == 0)
                                {
                                    add_flag = 1;
                                    best_cz.Add_Best(gk, v_send, t_all[i], E_all[i], S_error[i],travelTime,acc);
                                }
                                best_cz.Delete_Best(delete);
                            }
                            else if (flag_ori == 2)
                            {
                                Console.WriteLine("803 "+i.ToString()+" ");
                                best_cz.Add_Best(gk, v_send, t_all[i], E_all[i], S_error[i], travelTime,acc);
                                best = best.Next;
                            }
                        }
                    }
                }
            }
            best=best_cz.Return_Best();
            Console.WriteLine("number: "+best_cz.Count);
            for(int i = 0;i<qd_num;i++)
            {
                gk[i] = best.GK[i];
                v_send[i] = best.Speed[i];
            }
            v_send[qd_num] = 0;// best.GK[qd_num];
            //gk[qd_num] = -8;
            Best<int> best_jieji = best_cz.Head;
            int k1 = 0;
            while (best_jieji != null)
            {
                ///最优解集存储路径
                string best_txt = "E:\\中期\\粒子群\\单车\\" + DateTime.Now.Hour.ToString() +
                    DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString()+"num"+ k1.ToString() + ".txt";
                string wrote = "result:" + best_jieji.T_zu.ToString("0.00") + "s\ne:" + best_jieji.E.ToString("0.00") +
                "kJ\ns:" + best_jieji.S_ERROR.ToString("0.00") + "\n"; ;
                for (int i = 0; i < slope_ori.QD_NUM(); i++)
                {
                    //Console.WriteLine(gk[i].ToString() + "gk and speed" + speed[i].ToString() + " " + i.ToString());
                    wrote += best_jieji.GK[i].ToString() + " " + best_jieji.Speed[i].ToString() + " " + slope_ori.FindSlope(i).ToString() +" "+best_jieji.Time[i].ToString()+ "\n";
                }
                wrote +=  "0 0 0 " + best_jieji.T_zu.ToString("0.00") + "\n";
                File.WriteAllText(best_txt, wrote);
                k1++;
                best_jieji = best_jieji.Next;
            }
            return (gk, v_send, best.T_zu, best.E, best.S_ERROR);
        }
        public void Read_Excel(string file, string sheet_name, int fx)
        {
            string GetValueFromCell(DocumentFormat.OpenXml.Spreadsheet.Cell cell, SharedStringTable sharedStringTable)
            {
                if (cell.CellFormula != null)
                {
                    return cell.CellFormula.InnerText;
                }
                else
                {
                    if (cell.DataType == null)
                    {
                        return cell.InnerText;
                    }
                    else if (cell.DataType == CellValues.SharedString)
                    {
                        return sharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
                    }
                    else
                    {
                        return cell.InnerText;
                    }
                }
            }
            //读表
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, false)) //打开文件
            {
                //获取workbookPart对象
                WorkbookPart workbookPart = document.WorkbookPart;
                // 获取所有的工作表
                IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                // 查找Sheet
                Sheet sheet_Grid = sheets.FirstOrDefault(s => s.Name == sheet_name);
                ///分相区和限速表名称
                Sheet sheet_FenXiangQv = sheets.FirstOrDefault(s => s.Name == "分相区");
                Sheet sheet_SpeedLimit = sheets.FirstOrDefault(s => s.Name == "speedlimit");
                //link
                if (sheet_Grid != null)
                {
                    // 获取Sheet2的WorksheetPart
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet_Grid.Id);
                    Worksheet worksheet = worksheetPart.Worksheet;
                    SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                    SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

                    int grid = 0;
                    int startpos_grid = 0;
                    int endpos_grid = 0;
                    int length_grid = 0;
                    int grid_num = 0;

                    foreach (DocumentFormat.OpenXml.Spreadsheet.Row row in worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>())
                    {
                        //获取Excel的列头
                        if (row.RowIndex > 1)
                        {
                            // 解析每一列的数据
                            foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                            {
                                string value = GetValueFromCell(cell, sharedStringTable);
                                Console.WriteLine(value);
                                if (cell.CellReference.Value.Contains("A"))
                                {
                                    grid = (int)(System.Convert.ToDouble(value));
                                    //grid =int.Parse(value);
                                }
                                else if (cell.CellReference.Value.Contains("B"))
                                {
                                    startpos_grid = int.Parse(value);
                                }
                                else if (cell.CellReference.Value.Contains("C"))
                                {
                                    endpos_grid = int.Parse(value);
                                }
                                else if (cell.CellReference.Value.Contains("D"))
                                {
                                    if (fx == 1)
                                        length_grid = int.Parse(value);
                                    if (fx == 0)
                                        length_grid = endpos_grid - startpos_grid;
                                }
                            }
                            if (fx == 0)
                            {
                                if (length_grid > 600)
                                {
                                    int len_linshi = 600;
                                    int startpos_linshi = startpos_grid;
                                    //int startpos_linshi = startpos_grid;
                                    int endpos_linshi = 0;//startpos_linshi + len_linshi;
                                    while (length_grid / 600 >= 2)
                                    {
                                        length_grid -= len_linshi;
                                        endpos_linshi = startpos_linshi + len_linshi;
                                        slope_ori.Add(grid, startpos_linshi, endpos_linshi, len_linshi);
                                        startpos_linshi = endpos_linshi;
                                    }
                                    while (length_grid >= 800)
                                    {
                                        len_linshi = 400;
                                        length_grid -= len_linshi;
                                        endpos_linshi = startpos_linshi + len_linshi;
                                        slope_ori.Add(grid, startpos_linshi, endpos_linshi, len_linshi);
                                        startpos_linshi = endpos_linshi;
                                    }
                                    endpos_linshi = endpos_grid;
                                    len_linshi = endpos_linshi - startpos_linshi;
                                    slope_ori.Add(grid, startpos_linshi, endpos_linshi, len_linshi);
                                }
                                else
                                {
                                    slope_ori.Add(grid, startpos_grid, endpos_grid, length_grid);
                                }
                            }
                            if (fx == 1)
                            {
                                if (length_grid > 600)
                                {
                                    int len_linshi = 600;
                                    //int startpos_linshi = 0;
                                    int startpos_linshi = startpos_grid;
                                    int endpos_linshi = 0;//startpos_linshi + len_linshi;
                                    while (length_grid / 600 >= 2)
                                    {
                                        length_grid -= len_linshi;
                                        endpos_linshi = startpos_linshi - len_linshi;
                                        slope_ori.Add(grid, startpos_linshi, endpos_linshi, len_linshi);
                                        startpos_linshi = endpos_linshi;
                                    }
                                    while (length_grid >= 800)
                                    {
                                        len_linshi = 400;
                                        length_grid -= len_linshi;
                                        endpos_linshi = startpos_linshi - len_linshi;
                                        slope_ori.Add(grid, startpos_linshi, endpos_linshi, len_linshi);
                                        startpos_linshi = endpos_linshi;
                                    }
                                    endpos_linshi = endpos_grid;
                                    len_linshi = -endpos_linshi + startpos_linshi;
                                    slope_ori.Add(grid, startpos_linshi, endpos_linshi, len_linshi);
                                }
                                else
                                {
                                    slope_ori.Add(grid, startpos_grid, endpos_grid, length_grid);
                                }
                            }
                        }
                    }
                }
                //画轨道区段
                if (sheet_FenXiangQv != null)
                {
                    // 获取Sheet2的WorksheetPart
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet_FenXiangQv.Id);
                    Worksheet worksheet = worksheetPart.Worksheet;
                    SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                    SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

                    int startpos_fxq = 0;
                    int endpos_fxq = 0;
                    int length_fxq = 0;
                    int fxq_num = 0;

                    foreach (DocumentFormat.OpenXml.Spreadsheet.Row row in worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>())
                    {
                        //获取Excel的列头
                        if (row.RowIndex > 1)
                        {
                            // 解析每一列的数据
                            foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                            {
                                string value = GetValueFromCell(cell, sharedStringTable);

                                if (cell.CellReference.Value.Contains("A"))
                                {
                                    startpos_fxq = int.Parse(value);
                                }
                                else if (cell.CellReference.Value.Contains("B"))
                                {
                                    endpos_fxq = int.Parse(value);
                                }
                                else if (cell.CellReference.Value.Contains("C"))
                                {
                                    length_fxq = endpos_fxq - startpos_fxq;
                                    //length_fxq = int.Parse(value);
                                }
                            }
                            fxq_ori.Add(startpos_fxq, endpos_fxq, length_fxq);
                        }
                    }

                }
            }

            //PekingKalgan_Train.PrintTrain();
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
