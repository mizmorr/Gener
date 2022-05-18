using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;
using System.Text.RegularExpressions;


namespace GeneratorTask
{
    public class Word
    {
        public DocX doc;

        public Word(int var_num, string Name,string filepathe)
        {
            string fileAnswName = "dEGeneratAnsw";
            doc = DocX.Create(filepathe);
            if (Name!=null)
            doc.InsertParagraph(Name).Alignment = Alignment.right;
            Paragraph p1 = doc.InsertParagraph();
            doc.InsertParagraph();
            p1.AppendLine("Вариант " + var_num.ToString()).Alignment = Alignment.center;
            doc.InsertParagraph();
            doc.Save();
        }
        public Word(string filename) => doc = DocX.Load(filename);

    }
    interface CreateTask
    {
        void addlist();
        void Task(string name);
    }

    public class CreateTask1_1 : CreateTask
    {
        Random random = new Random();
        List<int> list = new List<int>();
        public void addlist()
        {
            int[] input = { 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            list.AddRange(input);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            int fst = list[random.Next(list.Count)];
            int presnd = list[random.Next(list.Count)];
            int snd = presnd == fst ? list[random.Next(list.Count)] : presnd;
            int trd = list.Find(x => x != fst && x != snd);
            string res = "1. Спортивный комментатор забыл счет баскетбольного матча, но помнит, что каждая команда набрала меньше 100 очков. Какова вероятность того, что, объявляя счет наугад, комментатор правильно назовет число очков, набранных первой командой, если ему подсказали, что это число: \nа) не содержит цифр " + fst.ToString() + " и " + snd.ToString() + "; \nб) содержит цифру " + trd.ToString() + "? ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_2 : CreateTask
    {
        Random random = new Random();
        List<string> list = new List<string>();
        public void addlist()
        {
            string[] input = { "России", "Украины", "Болгарии" };
            list.AddRange(input);
        }
        private int Parse(string s)
        {
            switch (s)
            {
                case "России":
                    return 6;
                case "Украины":
                    return 5;
                case "Болгарии":
                    return 4;
            }
            return 0;
        }
        public void Task(string name)
        {
            addlist();

            Word word = new Word(name);
            word.doc.InsertParagraph();
            string fst = list[random.Next(list.Count)];
            string snd = list.Find(x => x != fst);
            string pretrd = list[random.Next(list.Count)];
            string trd = pretrd == snd ? list[random.Next(list.Count)] : pretrd;
            string res = "2. В третий тур конкурса красоты прошли 6 участниц из России, 5 — из Украины и 4 — из Болгарии. Для представления участниц на сцену наугад приглашают 5 девушек. Найти вероятность того, что среди приглашенных:\nа) все девушки из " + fst + "; \nб) две девушки из " + snd + " и две — из " + trd + ". ";
            int f = Parse(fst); int s = Parse(snd); int t = Parse(trd);
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_3 : CreateTask
    {
        Random random = new Random();
        List<string> list = new List<string>();
        public void addlist()
        {
            string[] input = { "А — выпало два герба;", "В — выпали герб и решка;", "С — в первом бросании выпал герб, во втором бросании герб не выпал;", "D — ни разу не выпал герб;" };
            list.AddRange(input);

        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            string head = "3. Пусть событие Г\u1D62 — появление герба в i-ом бросании монеты, Р\u1D62 — появление решки в j-ом бросании монеты. Монету подбрасывают два раза. Постройте множество элементарных исходов и выявите состав подмножеств, соответствующих событиям: ";
            int rand = random.Next(0, 2);
            var newlist = new List<string>();
            switch (rand)
            {
                case 0:
                    newlist = (list.OrderBy(x => x.Length)).ToList();
                    break;
                case 1:
                    newlist = (list.OrderByDescending(x => x.Length)).ToList();
                    break;
                case 2:
                    newlist = (list.OrderBy(x => x.Length % 2 == 0)).ToList();
                    break;
            }
            string Tail = "\n" + newlist[0] + "\n" + newlist[1] + "\n" + newlist[2] + "\n" + newlist[3];
            string res = head + Tail;
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_4 : CreateTask
    {
        Random random = new Random();
        List<double> prob = new List<double>();
        List<string> mem = new List<string>();
        public void addlist()
        {
            double[] input1 = { 0.4, 0.5, 0.6, 0.7, 0.8, 0.9 };
            string[] input = { "вам удастся получить оба автографа;", "удастся получить хотя бы один автограф;", "не удастся получить автограф у польской певицы? " };
            mem.AddRange(input);
            prob.AddRange(input1);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double fst = prob[random.Next(prob.Count)];
            double presnd = prob[random.Next(prob.Count)];
            double snd = presnd == fst ? prob[random.Next(prob.Count)] : presnd;
            string head = "4. Российская певица дает автограф с вероятностью " + fst.ToString().Replace('.', ',') + ", а польская — с вероятностью " + snd.ToString().Replace('.', ',') + ". Какова вероятность того, что завтра после концерта с участием обеих звезд: ";
            int rand = random.Next(0, 2);
            var newlist = new List<string>();
            switch (rand)
            {
                case 0:
                    newlist = (mem.OrderBy(x => x.Length)).ToList();
                    break;
                case 1:
                    newlist = (mem.OrderByDescending(x => x.Length)).ToList();
                    break;
                case 2:
                    newlist = (mem.OrderBy(x => x.Length % 2 == 0)).ToList();
                    break;
            }
            string tail = "\na) " + newlist[0] + "\nб) " + newlist[1] + "\nв) " + newlist[2];
            string res = head + tail;
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_5 : CreateTask
    {
        Random random = new Random();
        List<double> prob = new List<double>();
        List<string> list = new List<string>();
        public void addlist()
        {
            list.Add("первой"); list.Add("второй");
            double[] input1 = { 0.4, 0.5, 0.6, 0.7, 0.8 };
            prob.AddRange(input1);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double fst = prob[random.Next(prob.Count)]; double presnd = prob[random.Next(prob.Count)]; double snd = presnd == fst ? prob[random.Next(prob.Count)] : presnd;
            string fststr = list[random.Next(list.Count)]; string sndstr = list.Find(x => x != fststr);
            string res = "5. Две россиянки участвуют в международном конкурсе по мировой экономике. Успешно пройти тур первая девушка может с вероятностью " + fst.ToString().Replace('.', ',') + ", вторая — " + snd.ToString().Replace('.', ',') + ". Вчера прошел третий, последний тур соревнований. Какова вероятность того, что у " + fststr + " участницы успешно пройденных туров больше, чем у " + sndstr + "? ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_6 : CreateTask
    {
        Random random = new Random();
        List<string> list = new List<string>();
        List<string> list2 = new List<string>();
        public void addlist()
        {
            string[] input = { "вторую", "третью", "четвертую" };
            list.AddRange(input);
            string[] i2 = { "дублем", "не дублем" };
            list2.AddRange(i2);
        }
        private int Parse(string s)
        {
            switch (s)
            {
                case "вторую":
                    return 2;
                case "третью":
                    return 3;
                case "четвертую":
                    return 4;
            }
            return 0;
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();

            string head = "6. Из полного набора костей домино (28) наугад извлечена кость. Найти вероятность того, что вторую наугад взятую кость можно приставить к первой, если первая оказалась: ";
            string fst = list2[random.Next(list2.Count)];
            string snd = list2.Find(x => x != fst);
            string r = list[random.Next(list.Count)]; int rint = Parse(r);
            string tail = "\na) " + fst + "\nб) " + snd;
            string res = head + tail;
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }

    public class CreateTask1_7 : CreateTask
    {
        Random random = new Random();
        List<double> prob = new List<double>();
        public void addlist()
        {
            double[] input1 = { 0.3, 0.4, 0.5, 0.6, 0.7, 0.8 };
            prob.AddRange(input1);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double fst = prob[random.Next(prob.Count)];
            double presnd = prob[random.Next(prob.Count)]; double snd = presnd == fst ? prob[random.Next(prob.Count)] : presnd;
            string res = "7. Три брата посеяли пшеницу, однако «...в долгом времени аль вскоре приключилось с ними горе: кто-то в поле стал ходить да пшеницу шевелить. Наконец они смекнули, чтоб стоять на карауле, хлеб ночами поберечь, злого вора подстеречь». В их деревне всем известно, что старший брат засыпает в дозоре с вероятностью  " + fst.ToString().Replace('.', ',') + ", средний — " + snd.ToString().Replace('.', ',') + ", а у младшего бессонница. Найти вероятность того, что в первую ночь удастся поймать вора, если очередность дежурства определяется жребием. ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_8 : CreateTask
    {
        Random random = new Random();
        List<double> prob = new List<double>();
        public void addlist()
        {
            double[] input1 = { 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8 };
            prob.AddRange(input1);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double fst = prob[random.Next(prob.Count)]; double presnd = prob[random.Next(prob.Count)]; double snd = presnd == fst ? prob[random.Next(prob.Count)] : presnd; double trd = prob.Find(x => x != fst && x != snd);
            double fst2 = prob[random.Next(prob.Count)]; double presnd2 = prob[random.Next(prob.Count)]; double snd2 = presnd2 == fst2 ? prob[random.Next(prob.Count)] : presnd; double trd2 = prob.Find(x => x != fst2 && x != snd2);
            string res = "8. Зритель с вероятностью " + fst.ToString().Replace('.', ',') + ", " + snd.ToString().Replace('.', ',') + " и " + trd.ToString().Replace('.', ',') + " соответственно может обратиться за билетом в одну из трех театральных касс Большого театра: в помещении театра, на Тверской и на станции метро «Пушкинская». Вероятность того, что к моменту прихода зрителя в кассе все билеты будут проданы, соответственно равна " + fst2.ToString().Replace('.', ',') + ", " + snd2.ToString().Replace('.', ',') + " и " + trd2.ToString().Replace('.', ',') + ". Поклонник Большого театра купил билет в одной из этих трех касс. Какова вероятность того, что эта касса на Тверской? ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }

    public class CreateTask1_9 : CreateTask
    {
        Random random = new Random();
        List<string> words = new List<string>();
        public void addlist()
        {
            string[] input = { "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь" };
            words.AddRange(input);
        }
        public int Parse(string s)
        {
            switch (s)
            {
                case "один":
                    return 1;
                case "два":
                    return 2;
                case "три":
                    return 3;
                case "четыре":
                    return 4;
                case "пять":
                    return 5;
                case "шесть":
                    return 6;
                case "семь":
                    return 7;
                case "восемь":
                    return 8;
            }
            return 0;
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            string wd = words[random.Next(words.Count)];
            int wdint = Parse(wd);
            string res = "9. В ящике имеется 5 синих и 50 красных шаров. Какова вероятность того, что при десяти независимых выборах с возвращением " + wd + " раз будет выниматься синий шар? ";
            res = (wd == words[0] || wd == words[1] || wd == words[2]) ? Regex.Replace(res, "раза", "раз") : res;
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_10 : CreateTask
    {
        Random random = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            int num = random.Next(65, 85);
            string res = "10. Вероятность переключения передач на каждом километре трассы равна 0,25. Найти вероятность того, что на 243 километровом участке этой трассы переключение передач произойдет: \n   а) " + num.ToString() + " раз; \n  б) не более " + num.ToString() + " раз. ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_11 : CreateTask
    {
        Random random = new Random();
        List<string> words = new List<string>();
        public void addlist()
        {
            string[] input = { "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь" };
            words.AddRange(input);
            string wd = words[random.Next(words.Count)];
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            CreateTask1_9 c = new CreateTask1_9();
            string wd = words[random.Next(words.Count)];
            int wdint = c.Parse(wd);
            string res = "11. Вероятность выхода из строя во время испытания на надежность любого из однотипных приборов равна 0,001. Найти вероятность того, что в партии из 100 приборов во время испытания выйдут из строя не более " + wd + " приборов. ";
            res = (wd == words[0]) ? Regex.Replace(res, "приборов", "прибор") : res;
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_12 : CreateTask
    {
        Random random = new Random();
        List<double> prob = new List<double>();
        public void addlist()
        {
            double[] input1 = { 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8 };
            prob.AddRange(input1);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double fst = prob[random.Next(prob.Count)];
            double presnd = prob[random.Next(prob.Count)];
            double snd = fst > presnd ? prob[random.Next(prob.Count)] : presnd;
            double pretrd = prob[random.Next(prob.Count)];
            double trd = snd > pretrd ? prob[random.Next(prob.Count)] : pretrd;
            string res = "12. Производится три независимых выстрела. Вероятность попадания при первом выстреле равна " + fst.ToString().Replace('.', ',') + "; при втором — " + snd.ToString().Replace('.', ',') + "; при третьем — " + trd.ToString().Replace('.', ',') + ". Составить ряд распределения числа попаданий. Найти М(Х), D(X),\u03C3(X), F(X) этой случайной величины.Построить график F(X). ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_13 : CreateTask
    {
        Random random = new Random();
        List<double> prob = new List<double>();
        public void addlist()
        {
            double[] input1 = { 0.3, 0.4, 0.5, 0.6, 0.7, 0.8 };
            prob.AddRange(input1);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double probability = prob[random.Next(prob.Count)];
            string res = "13. Радиосигнал передан четыре раза. Вероятность приема одного из них равна " + probability.ToString().Replace('.', ',') + ". Составить ряд распределения числа передач, в которых сигнал будет принят. Найти M(X) и D(X) этой случайной величины. ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_14 : CreateTask
    {
        Random random = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double prob = Math.Round(random.NextDouble(), 3);
            string res = "14. Вероятность выхода из строя электронной лампы, проработавшей t дней, равна " + prob.ToString().Replace('.', ',') + ". Аппаратура содержит 1000 ламп. Составить ряд распределения числа вышедших из строя ламп, проработавших t дней. Найти M(X) этой случайной величины. ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_15 : CreateTask
    {
        Random rand = new Random();
        List<double> prob = new List<double>();
        public void addlist()
        {
            double[] input = { 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7 };
            prob.AddRange(input);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            string hat = "15. Независимые случайные величины X и Y заданы таблицами распределений. \nНайти: \n1) M(X), M(Y), D(X), D(Y); \n2) таблицы распределения случайных величин Z\u2081 = = 2X+Y, Z\u2082 = X Y; \n3) M(Z\u2081), M(Z\u2082), D(Z\u2081), D(Z\u2082) непосредственно по таблицам распределений и на основании свойств математического ожидания и дисперсии. ";
            word.doc.InsertParagraph(hat);
            Table t = word.doc.AddTable(2, 4);
            float[] widt = { 30, 30, 30, 30 };
            t.SetWidths(widt);
            t.Rows[0].Cells[0].Paragraphs.First().Append("x\u1D62");
            t.Rows[1].Cells[0].Paragraphs.First().Append("p\u1D62");
            t.Rows[0].Cells[1].Paragraphs.First().Append("-1");
            t.Rows[0].Cells[2].Paragraphs.First().Append("1");
            t.Rows[0].Cells[3].Paragraphs.First().Append("2");

            string fst = prob[rand.Next(prob.Count)].ToString().Replace('.', ',');
            string presnd = prob[rand.Next(prob.Count)].ToString().Replace('.', ',');
            string snd = String.Equals(presnd, fst) ? prob[rand.Next(prob.Count)].ToString().Replace('.', ',') : presnd;
            t.Rows[1].Cells[1].Paragraphs.First().Append(fst);
            t.Rows[1].Cells[2].Paragraphs.First().Append(snd);
            t.Rows[1].Cells[3].Paragraphs.First().Append("p");
            word.doc.InsertParagraph("\n\t");
            word.doc.InsertTable(t).Alignment = Alignment.center;
            word.doc.InsertParagraph("\n\t");
            Table t2 = word.doc.AddTable(2, 3);
            float[] widt2 = { 30, 30, 30 };
            t2.SetWidths(widt2);
            t2.Rows[0].Cells[0].Paragraphs.First().Append("y\u1D62");
            t2.Rows[1].Cells[0].Paragraphs.First().Append("p\u1D62");
            t2.Rows[0].Cells[1].Paragraphs.First().Append("3");
            t2.Rows[0].Cells[2].Paragraphs.First().Append("5");
            string fst2 = prob[rand.Next(prob.Count)].ToString().Replace('.', ',');
            string presnd2 = prob[rand.Next(prob.Count)].ToString().Replace('.', ',');
            string snd2 = String.Equals(presnd2, fst2) ? prob[rand.Next(prob.Count)].ToString().Replace('.', ',') : presnd2;
            t2.Rows[1].Cells[1].Paragraphs.First().Append(fst2);
            t2.Rows[1].Cells[2].Paragraphs.First().Append(snd2);
            word.doc.InsertTable(t2).Alignment = Alignment.center;
            word.doc.Save();
        }
    }
    public class CreateTask1_16 : CreateTask
    {
        Random rand = new Random();
        List<double> prob = new List<double>();
        public void addlist()
        {
            double[] input = { 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9 };
            prob.AddRange(input);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            word.doc.InsertParagraph("16. Дана функция распределения F(x) непрерывной случайной величины X. \nТребуется: \n1) найти плотность вероятности f(x); \n2) построить графики F(x) и f(x); \n3) найти Р(\u03B1 < X < \u03B2 ) для данных \u03B1, \u03B2");
            word.doc.InsertParagraph();
            Paragraph paragraph = word.doc.InsertParagraph();
            Image image = word.doc.AddImage(@"Task1_16.png");
            Picture p = image.CreatePicture();
            string alpha = prob[rand.Next(prob.Count)].ToString().Replace('.', ',');
            string preb = prob[rand.Next(prob.Count)].ToString().Replace('.', ',');
            string b = String.Equals(preb, alpha) ? prob[rand.Next(prob.Count)].ToString().Replace('.', ',') : preb;
            paragraph.AppendPicture(p).Alignment = Alignment.left;
            word.doc.InsertParagraph("\u03B1=" + alpha + ";  \u03B2= " + b);
            word.doc.Save();
        }
    }
    public class CreateTask1_17 : CreateTask
    {
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            word.doc.InsertParagraph("17. Дана плотность вероятности f(x) непрерывной случайной величины X. \nТребуется: \n1) найти параметр \u03B1; \n2) найти функцию распределения F(x); \n3) построить графики f(x) и F(x); \n4) найти асимметрию и эксцесс X. ");
            word.doc.InsertParagraph();
            Paragraph paragraph = word.doc.InsertParagraph();
            Image image = word.doc.AddImage(@"Task1_17.png");
            Picture p = image.CreatePicture();
            paragraph.AppendPicture(p).Alignment = Alignment.left;
            word.doc.Save();
        }
    }
    public class CreateTask1_18 : CreateTask
    {
        Random rand = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            word.doc.InsertParagraph("18.");
            Paragraph p1 = word.doc.InsertParagraph();
            Image image = word.doc.AddImage(@"Task2_181.png");
            Picture p = image.CreatePicture();
            p1.AppendPicture(p).Alignment = Alignment.left;
            word.doc.InsertParagraph();
            Paragraph p2 = word.doc.InsertParagraph();
            Image image2 = word.doc.AddImage(@"Task1_18.png");
            Picture pi2 = image2.CreatePicture();
            p2.AppendPicture(pi2).Alignment = Alignment.left;
            int alpha = rand.Next(0, 4);
            int prebet = rand.Next(4, 7);
            int bet = alpha > prebet ? rand.Next(4, 7) : prebet;
            word.doc.InsertParagraph("\u03B1=" + alpha.ToString() + ";  \u03B2= " + bet.ToString());
            word.doc.Save();
        }
    }
    public class CreateTask1_19 : CreateTask
    {
        Random rand = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            int a = rand.Next(1, 15);
            int b = rand.Next(50, 65);
            string res = "19. Случайная величина X имеет нормальный закон распределения (MX = 50; DX = 250). Найти вероятность события {X\u2208 (" + a.ToString() + ", " + b.ToString() + ")}. ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_20 : CreateTask
    {
        Random random = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double prob = 1;
            double preprob = Math.Round(random.NextDouble(), 2);
            while (prob > 0.1 && prob == 0)
            { prob = preprob > 0.1 ? Math.Round(random.NextDouble(), 2) : preprob; }
            string res = "20. Цена деления шкалы амперметра равна 0,1 А. Показания определяют с точностью до ближайшего деления. Найти вероятность того, что при отсчете будет сделана ошибка , превышающая " + prob.ToString().Replace('.', ',') + " А. ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask1_21 : CreateTask
    {
        Random rand = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            int hours = rand.Next(1, 12);
            string res = "21. Период накопления состава на сортировочной станции имеет нормальное распределение с параметрами: m = 10 ч и \u03C3= 1,5 ч. С какой вероятностью период накопления очередного состава окажется более " + hours.ToString() + " ч? ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_1 : CreateTask
    {
        Random random = new Random();
        List<string> nums = new List<string>();
        public void addlist()
        {
            string[] input = { "первого", "второго", "третьего", "четвертого" };
            nums.AddRange(input);
        }
        private int Parser(string s)
        {
            switch (s)
            {
                case "первого":
                    return 1;
                case "второго":
                    return 2;
                case "третьего":
                    return 3;
                case "четвертого":
                    return 4;
            }
            return 0;
        }
        public virtual void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            int kurss = random.Next(nums.Count);
            int kurs2 = random.Next(nums.Count);
            string kurs2a = nums[kurs2]; int k2a = Parser(kurs2a);
            string kurs1_1 = nums[kurss]; int k1_1 = Parser(kurs1_1);
            string kurs1_2 = nums.Find(x => x != kurs1_1); int k1_2 = Parser(kurs1_2);
            string kurs1_3 = nums.Find(x => x != kurs1_1 && x != kurs1_2); int k1_3 = Parser(kurs1_3);
            word.doc.InsertParagraph("1. В финальном забеге на 100 м участвуют по два студента с четырех курсов.\nНайти вероятность того, что:\nа) первым пробежит дистанцию студент " + kurs1_1 + " курса, вторым — студент " + kurs1_2 + " курса и третьим — студент " + kurs1_3 + " курса;\nб) в тройке призеров не будет студентов " + kurs2a + " курса.").Alignment = Alignment.left;
            word.doc.InsertParagraph();
            word.doc.Save();
        }
    }
    public class CreateTask2_2 : CreateTask
    {
        Random random = new Random();
        List<string> nums = new List<string>();
        List<string> numsb = new List<string>();
        public void addlist()
        {
            string[] input = { "один", "два", "три" };
            nums.AddRange(input);
            string[] inputb = { "одно", "два", "три" };
            numsb.AddRange(inputb);
        }
        public void Task(string name)
        {
            addlist();
            int a = random.Next(nums.Count);
            int b1 = random.Next(numsb.Count);
            int b2 = random.Next(numsb.Count);
            string aword = nums[a];
            string b1word = numsb[b1];
            string b2word = numsb[b2];
            string hat = "2. В студенческой столовой на обед предлагается по три вида салатов, первых и вторых блюд. Студент, как обычно, берет на обед пять блюд. Найти вероятность того, что он взял:";
            string input = "\nа) " + aword + " салата;\nб) " + b1word + " первых и " + b2word + " вторых блюда.";
            if (aword == "один") input = (Regex.Replace(input, "салата", "салат")).ToString();
            if (b1word == numsb[0]) input = (Regex.Replace(input, "первых", "первое")).ToString();
            if (b2word == numsb[0]) input = (Regex.Replace(input, "вторых блюда", "второе блюдо")).ToString();
            CreateTask1_9 c = new CreateTask1_9();
            int aw = c.Parse(aword); int b1w = c.Parse(b1word); int b2w = c.Parse(b2word);
            string res = hat + input;
            Word word = new Word(name);
            word.doc.InsertParagraph();
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_3 : CreateTask
    {
        Random random = new Random();
        List<string> nums = new List<string>();
        public void addlist()
        {
            string[] input = { "A ⋂ B", "A ⋃ B", "C\u0305", "A \u005C B\u0305" };
            nums.AddRange(input);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            string a = nums[random.Next(nums.Count)];
            string b = nums.Find(x => x != a);
            string c = nums.Find(x => x != b && x != a);
            string d = nums.Find(x => x != a && x != b && x != c);
            string res = "3. Пусть Х — число очков, выпавших на верхней грани игральной кости при однократном бросании. Рассмотрим следующие события: А — Х кратно трем; В — Х нечетно, С — Х больше трех. Постройте множество элементарных исходов и выявите состав подмножеств, соответствующих событиям:\na) " + a + "\nб) " + b + "\nв) " + c + "\nг) " + d;

            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_4 : CreateTask
    {
        Random random = new Random();
        List<double> prob = new List<double>();
        List<string> rightornot = new List<string>();
        List<string> who = new List<string>();

        public void addlist()
        {
            double[] input1 = { 0.6, 0.7, 0.8, 0.9 };
            string[] input2 = { "правильно", "неправильно" };
            string[] input3 = { "первый", "второй" };
            prob.AddRange(input1);
            rightornot.AddRange(input2);
            who.AddRange(input3);

        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double fst = prob[random.Next(prob.Count)];
            double snd = prob.Find(x => x != fst);
            string whostr = who[random.Next(who.Count)];
            string rightornotstr = rightornot[random.Next(rightornot.Count)];
            word.doc.InsertParagraph("4. В поликлинике работают два психолога. Первый правильно определяет профессиональные наклонности детей с вероятностью " + fst.ToString().Replace('.', ',') + " , второй — с вероятностью " + snd.ToString().Replace('.', ',') + ". Для большей надежности мама с ребенком посетила обоих психологов. Какова вероятность того, что:\nа)  профессиональные наклонности ребенка оба специалиста определят " + rightornotstr + ";\nб)  хотя бы один из них ошибется; \nв)  ошибочные рекомендации даст " + whostr + " психолог?");
            word.doc.Save();
        }
    }
    public class CreateTask2_5 : CreateTask
    {
        Random random = new Random();
        List<double> prob = new List<double>();
        List<double> prob2 = new List<double>();
        public void addlist()
        {
            double[] input1 = { 0.75, 0.8, 0.85, 0.9, 0.95 };
            double[] input2 = { 0.6, 0.7, 0.8, 0.9 };
            prob.AddRange(input1);
            prob2.AddRange(input2);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double fst = prob[random.Next(prob.Count)];
            double presnd = prob[random.Next(prob.Count)];
            double snd = Equals(presnd, fst) ? prob[random.Next(prob.Count)] : presnd;
            double sfst = prob2[random.Next(prob2.Count)];
            double spresnd = prob2[random.Next(prob2.Count)];
            double ssnd = Equals(spresnd, sfst) ? prob2[random.Next(prob2.Count)] : spresnd;
            word.doc.InsertParagraph("5. Инженер-электронщик и киноартист пытаются пополнить ряды космонавтов. С вероятностью " + fst + " и " + snd + " соответственно они успешно проходят тест по специальности, с вероятностью " + sfst + " и " + ssnd + " — по физической подготовке. Какова вероятность того, что киноартист успешно пройдет тестов больше, чем инженер-электронщик?");
            word.doc.Save();
        }
    }
    public class CreateTask2_6 : CreateTask
    {
        Random random = new Random();
        List<int> nums = new List<int>();
        public void addlist()
        {
            int[] input = { 2, 3, 4, 5 };
            nums.AddRange(input);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            int num = nums[3];
            string res = "6. В колоде 36 карт. Наугад извлекают " + num.ToString() + " карты. Найти вероятность того, что вторым вынут туз, если первым тоже вынут туз. ";
            if (num == 5) res = Regex.Replace(res, "карты", "карт");
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_7 : CreateTask
    {
        Random random = new Random();
        List<double> nums = new List<double>();
        public void addlist()
        {
            double[] input = { 0.1, 0.2, 0.3, 0.4, 0.5 };
            nums.AddRange(input);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double fst = nums[random.Next(nums.Count)];
            double presnd = nums[random.Next(nums.Count)];
            double snd = Equals(presnd, fst) ? nums[random.Next(nums.Count)] : presnd;
            double thrd = nums.Find(x => x != fst && x != snd);
            string res = "7. В фотоателье работают три оператора, каждый из которых печатает соответственно 35, 40 и 25% всей продукции. Вероятность того, что фотография будет некачественной, для первого оператора равна " + fst.ToString().Replace('.', ',') + ", для второго — " + snd.ToString().Replace('.', ',') + ", для третьего — " + thrd.ToString().Replace('.', ',') + ". Вы не знаете, к какому из операторов попала ваша фотопленка с портретом любимой бабушки. Какова вероятность того, что вы, получив снимок, узнаете на нем свою бабушку? ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_8 : CreateTask
    {
        Random random = new Random();
        List<double> prob1 = new List<double>();
        List<double> prob2 = new List<double>();
        public void addlist()
        {
            double[] input2 = { 0.2, 0.3, 0.4, 0.5, 0.6, 0.7 };
            prob2.AddRange(input2);
            double[] input1 = { 0.1, 0.2, 0.3, 0.4, 0.5 };
            prob1.AddRange(input1);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double fst = prob1[random.Next(prob1.Count)];
            double presnd = prob1[random.Next(prob1.Count)];
            double snd = Equals(presnd, fst) ? prob1[random.Next(prob1.Count)] : presnd;
            double thrd = prob1.Find(x => x != fst && x != snd);

            double fst2 = prob2[random.Next(prob2.Count)];
            double presnd2 = prob2[random.Next(prob2.Count)];
            double snd2 = Equals(presnd2, fst2) ? prob2[random.Next(prob2.Count)] : presnd2;
            double thrd2 = prob2.Find(x => x != fst2 && x != snd2);
            string res = "8. Студента Зевского на лекциях по математике посещают музы: Евтерпа (муза лирической поэзии) — с вероятностью " + fst.ToString().Replace('.', ',') + "; Эрато (муза любовной поэзии) — с вероятностью " + snd.ToString().Replace('.', ',') + " и Каллиопа (муза эпической поэзии) — с вероятностью " + thrd.ToString().Replace('.', ',') + ". Известно, что после посещения соответствующей музы Зевский лирические стихи сочиняет с вероятностью " + fst2.ToString().Replace('.', ',') + ", любовные — с вероятностью " + snd2.ToString().Replace('.', ',') + " и эпические — с вероятностью " + thrd2.ToString().Replace('.', ',') + ". Какова вероятность того, что написанное Зевским на очередной лекции стихотворение было эпическим? ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_9 : CreateTask
    {
        Random random = new Random();
        List<string> prob = new List<string>();
        public void addlist()
        {
            string[] input = { "одного", "двух", "трех", "четырех" };
            prob.AddRange(input);
        }
        private int Parser(string s)
        {
            switch (s)
            {
                case "одного":
                    return 1;
                case "двух":
                    return 2;
                case "трех":
                    return 3;
                case "четырех":
                    return 4;

            }
            return 0;
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            string num = prob[random.Next(prob.Count)]; int n = Parser(num);
            string res = "9. Для стрелка, выполняющего упражнение в тире, вероятность попасть в «яблочко» при одном выстреле не зависит от результатов предшествующих выстрелов и равна 0,25. Спортсмен сделал пять выстрелов. Найти вероятность не менее " + num + " попаданий. ";
            if (num == prob[0]) res = Regex.Replace(num, "попаданий", "попадания");
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_10 : CreateTask
    {
        Random random = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            int fst = random.Next(60, 75);
            int snd = random.Next(80, 95);
            string res = "10. Фабрика выпускает в среднем 80% продукции первого сорта. Какова вероятность того, что в партии из 100 изделий окажется: \nа) не менее " + fst.ToString() + " и не более " + snd.ToString() + " изделий первого сорта;\nб) ровно половина таких изделий? ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_11 : CreateTask
    {
        Random rand = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            string fst = rand.Next(65, 85).ToString();
            if (fst == "71" || fst == "81") fst = Regex.Replace(fst, "студентов", "студента");
            string snd = rand.Next(3, 5).ToString();
            string res = "11. Известно, что в среднем 5% студентов носят очки. Какова вероятность того, что из " + fst + " студентов, сидящих в аудитории, окажутся " + snd + " пользующихся очками? ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_12 : CreateTask
    {
        Random rand = new Random();
        List<double> prob = new List<double>();
        public void addlist()
        {
            double[] input = { 0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4 };
            prob.AddRange(input);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double fst = prob[rand.Next(prob.Count)];
            double presnd = prob[rand.Next(prob.Count)];
            double snd = Equals(presnd, fst) ? prob[rand.Next(prob.Count)] : presnd;
            double thrd = prob.Find(x => x != fst && x != snd);
            double fth = prob.Find(x => x != snd && x != thrd && x != fst);
            word.doc.InsertParagraph("12. Дана система из четырех блоков (рис. 15). \nВ случае неисправности системы вероятность неисправности 1, 2, 3 и 4-го блоков равна " + fst.ToString().Replace('.', ',') + "; " + snd.ToString().Replace('.', ',') + "; " + thrd.ToString().Replace('.', ',') + " и " + fth.ToString().Replace('.', ',') + "\n ");
            Paragraph paragraph = word.doc.InsertParagraph();
            Image image = word.doc.AddImage(@"Task2_12.png");
            Picture p = image.CreatePicture();
            paragraph.AppendPicture(p).Alignment = Alignment.center;
            word.doc.InsertParagraph("соответственно, а время, необходимое для поиска неисправности в каждом блоке, — 5, 6, 10 и 9 мин. Одновременный выход из строя двух или более блоков считается невозможным. Составить ряд распределения для случайной величины Т — времени, необходимого для поиска неисправностей в системе. Найти М(T), D(T), \u03C3(T), F(T) этой случайной величины. Построить график F(T). ");
            word.doc.Save();
        }
    }
    public class CreateTask2_13 : CreateTask
    {
        Random rand = new Random();
        List<string> fst = new List<string>();
        List<string> snd = new List<string>();
        public void addlist()
        {
            string[] input = { "восемь", "девять", "десять", "одиннадцать" };
            fst.AddRange(input);
            string[] input2 = { "три", "четыре", "пять", "шесть", "семь" };
            snd.AddRange(input2);
        }
        private int Parser(string s)
        {
            switch (s)
            {
                case "три":
                    return 3;
                case "четыре":
                    return 4;
                case "пять":
                    return 5;
                case "шесть":
                    return 6;
                case "семь":
                    return 7;
                case "восемь":
                    return 8;
                case "девять":
                    return 9;
                case "десять":
                    return 10;
                case "одиннадцать":
                    return 11;
            }
            return 0;
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            string fststr = fst[rand.Next(fst.Count)]; int f = Parser(fststr);
            string sndstr = snd[rand.Next(snd.Count)]; int s = Parser(sndstr);
            string res = "13. Партия, насчитывающая 100 швейных машин, содержит " + fststr + " бракованных. Из всей партии с целью проверки качества случайным образом отбирается " + sndstr + " швейных машин. Составить ряд распределения числа бракованных машин среди отобранных. Найти M(X) и D(X) этой случайной величины. ";
            if (sndstr == "три" || sndstr == "четыре") res = Regex.Replace(res, "швейных машин", "швейные машины");
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_14 : CreateTask
    {
        Random rand = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            int num = rand.Next(250, 400);
            string res = "14. Книга содержит 400 страниц. Вероятность сделать опечатку на одной странице равна 0,0025. Составить ряд распределения числа опечаток на одной странице, если в книге их " + num.ToString() + ". Найти M(X) числа опечаток на одной странице. ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class CreateTask2_15 : CreateTask
    {
        Random rand = new Random();
        List<double> prob = new List<double>();
        public void addlist()
        {
            double[] input = { 0.1, 0.2, 0.3, 0.4 };
            prob.AddRange(input);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            string hat = "15. Независимые случайные величины X и Y заданы таблицами распределений. \nНайти: \n1) M(X), M(Y), D(X), D(Y); \n2) таблицы распределения случайных величин Z\u2081 = = 2X+Y, Z\u2082 = X Y; \n3) M(Z\u2081), M(Z\u2082), D(Z\u2081), D(Z\u2082) непосредственно по таблицам распределений и на основании свойств математического ожидания и дисперсии. ";
            word.doc.InsertParagraph(hat);
            Table t = word.doc.AddTable(2, 4);
            float[] widt = { 30, 30, 30, 30 };
            t.SetWidths(widt);
            t.Rows[0].Cells[0].Paragraphs.First().Append("x\u1D62");
            t.Rows[1].Cells[0].Paragraphs.First().Append("p\u1D62");
            t.Rows[0].Cells[1].Paragraphs.First().Append("1");
            t.Rows[0].Cells[2].Paragraphs.First().Append("2");
            t.Rows[0].Cells[3].Paragraphs.First().Append("3");
            t.Rows[1].Cells[1].Paragraphs.First().Append("p");
            double fst = prob[rand.Next(prob.Count)];
            double presnd = prob[rand.Next(prob.Count)];
            double snd = Equals(presnd, fst) ? prob[rand.Next(prob.Count)] : presnd;
            t.Rows[1].Cells[2].Paragraphs.First().Append(fst.ToString());
            t.Rows[1].Cells[3].Paragraphs.First().Append(snd.ToString());
            word.doc.InsertParagraph("\n\t");
            word.doc.InsertTable(t).Alignment = Alignment.center;
            word.doc.InsertParagraph("\n\t");
            Table t2 = word.doc.AddTable(2, 3);
            float[] widt2 = { 30, 30, 30 };
            t2.SetWidths(widt2);
            t2.Rows[0].Cells[0].Paragraphs.First().Append("y\u1D62");
            t2.Rows[1].Cells[0].Paragraphs.First().Append("p\u1D62");
            t2.Rows[0].Cells[1].Paragraphs.First().Append("-1");
            t2.Rows[0].Cells[2].Paragraphs.First().Append("4");
            double fst2 = prob[rand.Next(prob.Count)];
            double presnd2 = prob[rand.Next(prob.Count)];
            double snd2 = Equals(presnd2, fst2) ? prob[rand.Next(prob.Count)] : presnd2;
            t2.Rows[1].Cells[1].Paragraphs.First().Append(fst2.ToString());
            t2.Rows[1].Cells[2].Paragraphs.First().Append(snd2.ToString());
            word.doc.InsertTable(t2).Alignment = Alignment.center;
            word.doc.Save();
        }
    }
    public class CreateTask2_16 : CreateTask
    {
        Random rand = new Random();
        List<double> prob = new List<double>();
        public void addlist()
        {
            double[] input = { 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7 };
            prob.AddRange(input);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            word.doc.InsertParagraph("16. Дана функция распределения F(x) непрерывной случайной величины X. \nТребуется: \n1) найти плотность вероятности f(x); \n2) построить графики F(x) и f(x); \n3) найти Р(\u03B1 < X < \u03B2 ) для данных \u03B1, \u03B2");
            word.doc.InsertParagraph();
            Paragraph paragraph = word.doc.InsertParagraph();
            Image image = word.doc.AddImage(@"Task2_16.png");
            Picture p = image.CreatePicture();
            double alpha = prob[rand.Next(prob.Count)];
            double preb = prob[rand.Next(prob.Count)];
            double b = Equals(preb, alpha) ? prob[rand.Next(prob.Count)] : preb;
            paragraph.AppendPicture(p).Alignment = Alignment.left;
            word.doc.InsertParagraph("\u03B1=" + alpha.ToString() + ";  \u03B2= " + b.ToString());
            word.doc.Save();

        }
    }
    public class CreateTask2_17 : CreateTask
    {
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            word.doc.InsertParagraph("17. Дана плотность вероятности f(x) непрерывной случайной величины X. \nТребуется: \n1) найти параметр \u03B1; \n2) найти функцию распределения F(x); \n3) построить графики f(x) и F(x); \n4) найти асимметрию и эксцесс X. ");
            word.doc.InsertParagraph();
            Paragraph paragraph = word.doc.InsertParagraph();
            Image image = word.doc.AddImage(@"Task2_17.png");
            Picture p = image.CreatePicture();
            paragraph.AppendPicture(p).Alignment = Alignment.left;
            word.doc.Save();
        }
    }
    public class CreateTask2_18 : CreateTask
    {
        Random rand = new Random();
        List<double> prob = new List<double>();
        public void addlist()
        {
            double[] input = { 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5 };
            prob.AddRange(input);
        }
        public void Task(string name)
        {
            addlist();
            Word word = new Word(name);
            word.doc.InsertParagraph();
            word.doc.InsertParagraph("18.");
            Paragraph p1 = word.doc.InsertParagraph();
            Image image = word.doc.AddImage(@"Task2_181.png");
            Picture p = image.CreatePicture();
            p1.AppendPicture(p).Alignment = Alignment.left;
            word.doc.InsertParagraph();
            Paragraph p2 = word.doc.InsertParagraph();
            Image image2 = word.doc.AddImage(@"Task2_182.png");
            Picture pi2 = image2.CreatePicture();
            p2.AppendPicture(pi2).Alignment = Alignment.left;
            string alpha = rand.Next(1, 6).ToString();
            double bet = prob[rand.Next(prob.Count)];
            word.doc.InsertParagraph("\u03B1=" + alpha + ";  \u03B2= " + bet.ToString());
            word.doc.Save();
        }
    }

    public class CreateTask2_19 : CreateTask
    {
        Random rand = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            string hat = "19. Время T безотказной работы телевизора распределено по показательному закону с плотностью: ";
            word.doc.InsertParagraph(hat);
            int hours = rand.Next(850, 1000);
            Paragraph p1 = word.doc.InsertParagraph();
            Image image = word.doc.AddImage(@"Task2_19.png");
            Picture p = image.CreatePicture();
            p1.AppendPicture(p).Alignment = Alignment.left;
            string tail = "Найти вероятность того, что телевизор проработает без отказа не менее " + hours.ToString() + " ч. ";
            word.doc.InsertParagraph(tail);
            word.doc.Save();
        }
    }
    public class CreateTask2_20 : CreateTask
    {
        Random rand = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            double preprob = rand.NextDouble();
            double prob = 1;
            while (prob < 0.51 && prob == 0)
            {
                preprob = rand.NextDouble();
                prob = preprob < 0.5 ? rand.NextDouble() : preprob;
            }
            prob = Math.Round(prob, 3);
            string res = "20. Станок-автомат изготавливает ролики, контролируя их диаметр D. Считая, что величина D распределена нормально (m = 5 см; = 2 мм), найти интервал, в который с вероятностью " + prob.ToString() + " попадут диаметры роликов. ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }

    public class CreateTask2_21 : CreateTask
    {
        Random rand = new Random();
        public void addlist() { }
        public void Task(string name)
        {
            Word word = new Word(name);
            word.doc.InsertParagraph();
            int num = rand.Next(85, 100);
            string res = "21. Число вагонов в прибывающем на расформирование составе — нормальная случайная величина с математическим ожиданием m = 80 и = 6. Какова вероятность того, что в очередном составе будет не менее " + num.ToString() + " вагонов? ";
            word.doc.InsertParagraph(res);
            word.doc.Save();
        }
    }
    public class Creator
    {
        private void ConvertNumIntoTask(string name, int num,int type)
        {
            switch (num)
            {
                case 1:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_1 create11 = new CreateTask1_1();
                            create11.Task(name);
                            break;
                        case 2:
                            CreateTask2_1 create21 = new CreateTask2_1();
                            create21.Task(name);
                            break;
                    }
                    break;
                case 2:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_2 create12 = new CreateTask1_2();
                            create12.Task(name);
                            break;
                        case 2:
                            CreateTask2_2 create22 = new CreateTask2_2();
                            create22.Task(name);
                            break;
                    }
                    break;
                case 3:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_3 create13 = new CreateTask1_3();
                            create13.Task(name);
                            break;
                        case 2:
                            CreateTask2_3 create23 = new CreateTask2_3();
                            create23.Task(name);
                            break;
                    }
                    break;
                case 4:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_4 create14 = new CreateTask1_4();
                            create14.Task(name);
                            break;
                        case 2:
                            CreateTask2_4 create24 = new CreateTask2_4();
                            create24.Task(name);
                            break;
                    }
                    break;
                case 5:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_5 create15 = new CreateTask1_5();
                            create15.Task(name);
                            break;
                        case 2:
                            CreateTask2_5 create25 = new CreateTask2_5();
                            create25.Task(name);
                            break;
                    }
                    break;
                case 6:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_6 create16 = new CreateTask1_6();
                            create16.Task(name);
                            break;
                        case 2:
                            CreateTask2_6 create26 = new CreateTask2_6();
                            create26.Task(name);
                            break;
                    }
                    break;
                case 7:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_7 create17 = new CreateTask1_7();
                            create17.Task(name);
                            break;
                        case 2:
                            CreateTask2_7 create27= new CreateTask2_7();
                            create27.Task(name);
                            break;
                    }
                    break;
                case 8:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_8 create18 = new CreateTask1_8();
                            create18.Task(name);
                            break;
                        case 2:
                            CreateTask2_8 create28 = new CreateTask2_8();
                            create28.Task(name);
                            break;
                    }
                    break;
                case 9:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_9 create19 = new CreateTask1_9();
                            create19.Task(name);
                            break;
                        case 2:
                            CreateTask2_9 create29 = new CreateTask2_9();
                            create29.Task(name);
                            break;
                    }
                    break;
                case 10:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_10 create110 = new CreateTask1_10();
                            create110.Task(name);
                            break;
                        case 2:
                            CreateTask2_10 create210 = new CreateTask2_10();
                            create210.Task(name);
                            break;
                    }
                    break;
                case 11:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_11 create111 = new CreateTask1_11();
                            create111.Task(name);
                            break;
                        case 2:
                            CreateTask2_11 create211 = new CreateTask2_11();
                            create211.Task(name);
                            break;
                    }
                    break;
                case 12:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_12 create112 = new CreateTask1_12();
                            create112.Task(name);
                            break;
                        case 2:
                            CreateTask2_12 create212 = new CreateTask2_12();
                            create212.Task(name);
                            break;
                    }
                    break;
                case 13:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_13 create113 = new CreateTask1_13();
                            create113.Task(name);
                            break;
                        case 2:
                            CreateTask2_13 create213 = new CreateTask2_13();
                            create213.Task(name);
                            break;
                    }
                    break;
                case 14:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_14 create114 = new CreateTask1_14();
                            create114.Task(name);
                            break;
                        case 2:
                            CreateTask2_14 create214 = new CreateTask2_14();
                            create214.Task(name);
                            break;
                    }
                    break;
                case 15:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_15 create115 = new CreateTask1_15();
                            create115.Task(name);
                            break;
                        case 2:
                            CreateTask2_15 create215 = new CreateTask2_15();
                            create215.Task(name);
                            break;
                    }
                    break;
                case 16:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_16 create116 = new CreateTask1_16();
                            create116.Task(name);
                            break;
                        case 2:
                            CreateTask2_16 create216 = new CreateTask2_16();
                            create216.Task(name);
                            break;
                    }
                    break;
                case 17:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_17 create117 = new CreateTask1_17();
                            create117.Task(name);
                            break;
                        case 2:
                            CreateTask2_17 create217 = new CreateTask2_17();
                            create217.Task(name);
                            break;
                    }
                    break;
                case 18:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_18 create118 = new CreateTask1_18();
                            create118.Task(name);
                            break;
                        case 2:
                            CreateTask2_18 create218 = new CreateTask2_18();
                            create218.Task(name);
                            break;
                    }
                    break;
                case 19:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_19 create119 = new CreateTask1_19();
                            create119.Task(name);
                            break;
                        case 2:
                            CreateTask2_19 create219 = new CreateTask2_19();
                            create219.Task(name);
                            break;
                    }
                    break;
                case 20:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_20 create120 = new CreateTask1_20();
                            create120.Task(name);
                            break;
                        case 2:
                            CreateTask2_20 create220 = new CreateTask2_20();
                            create220.Task(name);
                            break;
                    }
                    break;
                case 21:
                    switch (type)
                    {
                        case 1:
                            CreateTask1_21 create121 = new CreateTask1_21();
                            create121.Task(name);
                            break;
                        case 2:
                            CreateTask2_21 create221 = new CreateTask2_21();
                            create221.Task(name);
                            break;
                    }
                    break;
            }
        }

        public void Create(string filepathe, int type, List<int> nums, string StudentName,int var)
        {
            Word word = new Word(var, StudentName, filepathe);
            foreach (var item in nums)
                ConvertNumIntoTask(filepathe, item, type);
        }
        public string ReturnTask(string filepathe)
        {
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = application.Documents.Open(filepathe);
            return doc.Content.Text;
        }
    }
    internal class Program
    {
        static void Main(string[] args)
        {
        }
    }
}