public class Run {
    public static void main(String[] args) {
        CellList cl = new CellList();
        ExcellMethods em = new ExcellMethods("C:/Users/dbsrl/OneDrive/바탕 화면/test/test.xlsx",cl);
        em.ReadandLoadExcell();
        cl.SortCells();
        em.WriteSortResult();
        em.PointYellow("경북대학교");
        //cl.printSheet();
    }
}
