report 88000 "Custom Email"
{
    DefaultLayout = Word;
    WordLayout = 'CustomEmailBody.docx';
    dataset
    {
        dataitem(Integer; Integer)
        {
            MaxIteration = 1;
            column(DateToday; Format(Today(), 0, '<Day,2>/<Month,2>/<Year4>'))

            {

            }
        }
    }
    requestpage
    {
        layout
        {
            area(content)
            {
                group(GroupName)
                {
                }
            }
        }
        actions
        {
            area(processing)
            {
            }
        }
    }
    var
  

}
