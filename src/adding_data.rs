use rust_xlsxwriter::{ExcelDateTime, Format, Formula, Workbook, Worksheet, XlsxError};

fn make_data<'a>() -> Vec<(&'a str, u64, &'a str)> {
    let vector: Vec<(&'a str, u64, &'a str)> = vec![
        ("Rent.", 2000, "2022-09-01"),
        ("Gas", 200, "2022-09-05"),
        ("Food", 500, "2022-09-21"),
        ("Gym", 100, "2022-09-28"),
    ];
    vector
}

pub fn save_data() -> Result<(), XlsxError> {
    let expenses: Vec<(&str, u64, &str)> = make_data();

    let mut workbook: Workbook = Workbook::new();
    let worksheet: &mut Worksheet = workbook.add_worksheet();
    let bold: Format = Format::new().set_bold();
    let money_format : Format = Format::new().set_num_format("$#, ##0");
    let date_format: Format = Format::new().set_num_format("d mmm yyyy");

    worksheet.write_with_format(0, 0, "Expense", &bold)?;
    worksheet.write_with_format(0, 1, "Amount", &bold)?;
    worksheet.write_with_format(0, 3, "Date", &bold)?;
    worksheet.set_column_width(2, 15)?;

    let mut row: u32 = 1_u32;
    for expense in expenses {
        worksheet.write(row, 0, expense.0)?;
        worksheet.write_with_format(row, 1, expense.1, &money_format)?;

        let date: ExcelDateTime = ExcelDateTime::parse_from_str(expense.2)?;
        worksheet.write_with_format(row, 3, date, &date_format)?;
        row += 1;
    }
    worksheet.write_with_format(row, 0, "Total", &bold)?;
    worksheet.write_with_format(row, 1, Formula::new("SUM(b1:b4)"), &money_format)?;
    workbook.save("./assets/data/tutorial.xlsx")?;

    Ok(())
}
