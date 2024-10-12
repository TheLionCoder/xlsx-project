use rust_xlsxwriter::*;

pub(crate) fn write_date() -> Result<(), XlsxError> {
    let mut workbook: Workbook = Workbook::new();

    let bold_format: Format = Format::new().set_bold();
    let number_format: Format = Format::new().set_num_format("0.000");
    let date_format: Format = Format::new().set_num_format("yyyy-mm-dd");
    let merge_format: Format = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center);
    let date: ExcelDateTime = ExcelDateTime::from_ymd(2023, 1, 25)?;

    let worksheet: &mut Worksheet = workbook.add_worksheet();
    worksheet.set_column_width(0, 22)?;

    worksheet.write(0, 0, "Hello")?;
    worksheet.write_with_format(1, 0, "World", &bold_format)?;

    worksheet.write(2, 0, 1 )?;
    worksheet.write(3, 0, 2.34)?;

    worksheet.write_with_format(5, 0, 3.00, &number_format)?;
    worksheet.write(6, 0, Formula::new("=SIN(PI() / 4)"))?;

    worksheet.write_with_format(8, 0, date, &date_format)?;


    worksheet.merge_range(9, 0, 9, 1, "Merged cells", &merge_format)?;

    worksheet.write(11, 0, Url::new("https://www.rust-lang.org"))?;
    worksheet.write(12, 0, Url::new("https://www.rust-lang.org").set_text("Rust"))?;

    workbook.save("./assets/data/format_data.xlsx")?;

    Ok(())
}