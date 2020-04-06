import java.util.ArrayList;
import java.util.Collections;

public class CellList {
    private ArrayList<Cell> cellList;
    public CellList(){
        this.cellList =new ArrayList<Cell>();
    }
    public void pushCell(String name , double ratio){
        Cell input = new Cell(name, ratio);
        this.cellList.add(input);
    }
    public void SortCells(){
        Collections.sort(cellList);
    }
    public ArrayList<Cell> getCellList(){
        return this.cellList;
    }
    public void printSheet() {
        for(int i =0; i< cellList.size() ;i++) {
            System.out.println("기관명 :"+ cellList.get(i).getname() +" 간접비율 "+ cellList.get(i).getRatio()+"%");
        }
    }
}
