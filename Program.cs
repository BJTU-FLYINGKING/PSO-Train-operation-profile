using NPOI.OpenXmlFormats.Dml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using static liziqun0117.Form1;

namespace liziqun0117
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
    public class Best<T>
    {
        public int[] GK { get; set; }     //加速度序列
        public double[] Speed { get; set; }
        public double T_zu { get; set; }      //最优时间
        public double E {  get; set; }        //能耗
        public double S_ERROR {  get; set; }
        public double[] Time { get; set; }
        public double[] Acc { get; set; }
        public Best<T> Next{ get; set; }     
        public Best(int[] gk, double[] speed,double zu, double e,double s,double[] time,double[] acc)
        {
            GK = gk;
            Speed = speed;
            T_zu = zu;
            E = e;
            S_ERROR = s;
            Time = time;
            Acc = acc;
            Next = null;
        }
    }
    public class Best_Set<T>
    {
        private Best<T> head;
        public int Count=1;
        public Best<T> Head
        {
            get
            {
                return head;
            }
            set
            {
                head = value;
            }
        }
        public Best_Set()
        {
            head = null;
        }
        public void Add_Best(int[] gk, double[] speed,double t, double e,double s,double[] time,double[] acc)
        {
            int[] gkCopy = new int[gk.Length];
            double[] speedCopy = new double[speed.Length];
            double[] timeCopy = new double[time.Length];
            double[] accCopy = new double[acc.Length];
            Array.Copy(gk, gkCopy, gk.Length);
            Array.Copy(speed, speedCopy, speed.Length);
            Array.Copy(time, timeCopy, time.Length);
            Array.Copy(acc, accCopy, acc.Length);

            Best<T> new_node=new Best<T>(gkCopy,speedCopy,t,e,s,timeCopy,accCopy);
            if(head==null)
            {
                head = new_node;
                return;
            }
            if(head!=null)
            {
                Best<T> temp = head;
                while(temp.Next!=null)
                {
                    temp=temp.Next;
                }
                temp.Next = new_node;
            }
            Count++;
        }
        public void Delete_Best(Best<T> delete)
        {
            Best<T> temp=head;
            if(delete==head)
            {
                head=delete.Next;
                //delete.Next=null;
                delete = null;
            }
            else if(delete!=head)
            {
                while(temp.Next!=delete) 
                {
                    temp=temp.Next;
                }
                temp.Next=delete.Next;
                //delete = null;
            }
            Count--;
        }
        public Best<T> Return_Best()
        {
            int seed = 0;
            byte[] bytes = new byte[4];
            //创建加密服务，实现加密随机数生成器
            System.Security.Cryptography.RNGCryptoServiceProvider rng = new System.Security.Cryptography.RNGCryptoServiceProvider();
            //加密数据存入字节数组
            rng.GetBytes(bytes);
            //转成整型数据返回，作为随机数生成种子
            seed = BitConverter.ToInt32(bytes, 0);
            Random ran = new Random(seed);
            Best<T> temp = head;
            int k = ran.Next(0, Count);
            for(int i = 0; i < k; i++)
            {
                temp= temp.Next;
            }
            return temp;
        }
    }
    public class SLOPE<T>
    {
        public int Grid { get; set; }
        public int Start_pos { get; set; }
        public int End_Pos { get; set; }
        public int Length { get; set; }
        public SLOPE<T> Next { get; set; }

        // 构造函数
        public SLOPE(int grid, int start_pos, int end_pos, int length)
        {
            Grid = grid;
            Start_pos = start_pos;
            End_Pos = end_pos;
            Length = length;
            Next = null;
        }


    }
    public class LinkList_SLOPE<T>
    {
        private SLOPE<T> head;//存储一个头结点
        public SLOPE<T> Head
        {
            get
            {
                return head;
            }
            set
            {
                head = value;
            }
        }

        public LinkList_SLOPE()
        {
            head = null;
        }
        public void Add(int grid, int start_pos, int end_pos, int length)
        {
            SLOPE<T> newNode = new SLOPE<T>(grid,start_pos,end_pos,length);//根据新的数据创建一个新的节点
            //如果头结点为空，那么这个新的节点就是头节点
            if (head == null)
            {
                head = newNode;
                return;
            }
            else
            {//把新来的结点放到 链表的尾部
             //要访问到链表的尾结点
                SLOPE<T> temp = head;
                while (temp.Next != null)
                {
                    temp = temp.Next;
                }
                temp.Next = newNode;//把新来的结点放到 链表的尾部
            }

        }
        public SLOPE<T> Find(int number)
        {
            // 遍历链表中的每一个节点，查找编号为 jinluID 的节点
            SLOPE<T> current = head;
            int i = 0;
            while (i < number)
            {
                i++;
                current = current.Next;
            }
            return current; // 没有找到目标节点，返回 null
        }
        public int FindLength(int number)
        {
            // 遍历链表中的每一个节点，查找编号为 jinluID 的节点
            SLOPE<T> current = head;
            int i = 0;
            while (i<number) 
            {
                i++;
                current = current.Next;
            }
            return current.Length; // 没有找到目标节点，返回 null
        }
        public int FindSlope(int number)
        {
            // 遍历链表中的每一个节点，查找编号为 jinluID 的节点
            SLOPE<T> current = head;
            int i = 0;
            while (i < number)
            {
                i++;
                current = current.Next;
            }
            return current.Grid; // 没有找到目标节点，返回 null
        }
        public int QD_NUM()
        {
            SLOPE<T> current = head;
            int qd_num = 0;
            while(current != null)
            {
                qd_num++;
                current = current.Next;
            }
            return qd_num;
        }

        public int DISTANCE()
        {
            SLOPE<T> current = head;
            int qd_length = 0;
            while (current != null)
            {
                qd_length+=current.Length;
                current = current.Next;
            }
            return qd_length;
        }
        /// 制动点位置判断
        public int BREAK_POS(double length)  
        {
            SLOPE<T> current = head;
            int qd_num = 0;
            while (current != null)
            {
                if((current.Start_pos<=length)&&(current.End_Pos>=length))
                {
                    break;
                }
                qd_num++;
                current = current.Next; 
            }
            return qd_num;
        }
        public void QD_LAST(int fx)
        {
            SLOPE<T> current = head;
            int qd_num = 0;
            while (current.Next != null)
            {
                current = current.Next;
            }
            SLOPE<T> delete = current;
            int length = current.Length;
            int len_linshi = 100;
            int len = 100;
            int startpos_linshi = current.Start_pos;
            int end = current.End_Pos;
            int endpos_linshi = current.End_Pos;
            if (fx == 0)
            {
                while (length > len)
                {
                    len_linshi = len;
                    length -= len_linshi;
                    endpos_linshi = startpos_linshi + len_linshi;
                    SLOPE<T> newNode = new SLOPE<T>(current.Grid, startpos_linshi, endpos_linshi, len_linshi);//根据新的数据创建一个新的节点
                    current.Next = newNode;
                    current = current.Next;
                    startpos_linshi = endpos_linshi;
                }
                endpos_linshi = end;
                len_linshi = endpos_linshi - startpos_linshi;
                slope_ori.Add(current.Grid, startpos_linshi, endpos_linshi, len_linshi);
            }
            if (fx == 1)
            {
                while (length > len)
                {
                    len_linshi = len;
                    length -= len_linshi;
                    endpos_linshi = startpos_linshi - len_linshi;
                    SLOPE<T> newNode = new SLOPE<T>(current.Grid, startpos_linshi, endpos_linshi, len_linshi);//根据新的数据创建一个新的节点
                    current.Next = newNode;
                    current = current.Next;
                    startpos_linshi = endpos_linshi;
                }
                endpos_linshi = end;
                len_linshi = -endpos_linshi + startpos_linshi;
                slope_ori.Add(current.Grid, startpos_linshi, endpos_linshi, len_linshi);
            }
            SLOPE<T> current1 = head;
            while (current1.Next != delete)
            {
                current1 = current1.Next;
            }
            current1.Next = delete.Next;
        }
    }
    public class FENXIANGQV<T>
    {
        public int Start_pos { get; set; }
        public int End_Pos { get; set; }
        public int Length { get; set; }
        public FENXIANGQV<T> Next { get; set; }

        // 构造函数
        public FENXIANGQV(int start_pos, int end_pos, int length)
        {
            Start_pos = start_pos;
            End_Pos = end_pos;
            Length = length;
            Next = null;
        }


    }
    public class LinkList_FENXIANGQV<T>
    {
        private FENXIANGQV<T> head;//存储一个头结点
        public FENXIANGQV<T> Head
        {
            get
            {
                return head;
            }
            set
            {
                head = value;
            }
        }

        public LinkList_FENXIANGQV()
        {
            head = null;
        }
        public void Add(int start_pos, int end_pos, int length)
        {
            FENXIANGQV<T> newNode = new FENXIANGQV<T>(start_pos, end_pos, length);//根据新的数据创建一个新的节点
                                                                              //如果头结点为空，那么这个新的节点就是头节点
            if (head == null)
            {
                head = newNode;
                return;
            }
            else
            {//把新来的结点放到 链表的尾部
             //要访问到链表的尾结点
               FENXIANGQV<T> temp = head;
                while (temp.Next != null)
                {
                    temp = temp.Next;
                }
                temp.Next = newNode;//把新来的结点放到 链表的尾部
            }

        }
        public FENXIANGQV<T> Find(int number)
        {
            // 遍历链表中的每一个节点，查找编号为 jinluID 的节点
            FENXIANGQV<T> current = head;
            int i = 0;
            while (i<number)
            {
                i++;
                current = current.Next;
            }
            return current; // 没有找到目标节点，返回 null
        }
    }
}
