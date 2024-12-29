#include <QFileDialog>
#include <QMessageBox>
#include <QStandardPaths>
#include <QTableView>
#include <QWidget>

#include "../QXlsx/header/xlsxdocument.h"

int TableExcelByXlsx(QWidget *ui, QTableView *tableView, const QString &title)
{
    QFileDialog dialog;
    QString     fileName
        = dialog.getSaveFileName(tableView, "Save Excel File", QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation), "Excel Files (*.xlsx)");
    if (!fileName.isEmpty()) {
        QXlsx::Document xlsx;

        int colCount = tableView->model()->columnCount();
        int rowCount = tableView->model()->rowCount();

        // 标题行
        xlsx.write(1, 1, title);
        QXlsx::Format titleFormat;
        titleFormat.setFontSize(18);
        titleFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
        xlsx.mergeCells(QXlsx::CellRange(1, 1, 1, colCount), titleFormat);
        xlsx.setRowHeight(1, 30);

        // 列标题
        QXlsx::Format headerFormat;
        headerFormat.setFontBold(true);
        headerFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
        headerFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
        headerFormat.setPatternBackgroundColor(QColor(191, 191, 191));
        for (int i = 0; i < colCount; i++) {
            QString columnName = tableView->model()->headerData(i, Qt::Horizontal, Qt::DisplayRole).toString();
            xlsx.write(2, i + 1, columnName, headerFormat);
            xlsx.setColumnWidth(i + 1, tableView->columnWidth(i) / 6);
        }
        xlsx.setRowHeight(2, 20);

        // 数据行
        QXlsx::Format dataFormat;
        dataFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
        dataFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
        for (int i = 0; i < rowCount; i++) {
            for (int j = 0; j < colCount; j++) {
                QModelIndex index   = tableView->model()->index(i, j);
                QString     strData = tableView->model()->data(index).toString();
                xlsx.write(i + 3, j + 1, strData, dataFormat);
            }
        }
        xlsx.setRowHeight(3, 20);

        // 保存文件
        xlsx.saveAs(fileName);

        int result = QMessageBox::question(ui, "Export Finished", "Do you want to open the exported Excel file?", QMessageBox::Yes | QMessageBox::No);
        if (result == QMessageBox::Yes) {
            QDesktopServices::openUrl(QUrl::fromLocalFile(fileName));
        }
        return 0;
    } else {
        QMessageBox::warning(ui, "Warning", "Failed to save the Excel file.");
        return -1;
    }
}
