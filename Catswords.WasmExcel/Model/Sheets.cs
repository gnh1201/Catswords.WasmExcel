namespace Catswords.WasmExcel.Model
{
    public class Sheets
    {
        public Books Parent { get; set; }

        public Sheets(Books parent)
        {
            Parent = parent;
        }
    }
}
