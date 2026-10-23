<?php

declare(strict_types=1);

namespace SilverShop\Import\Excel\Export;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Formula-injection-safe cell writing for the export workbooks. A value such as {@code =1+1},
 * {@code +...}, {@code -...}, {@code @...} (or a leading tab / CR) is a spreadsheet-formula trigger: a
 * client opening our export in Excel / LibreOffice would see it evaluated — and a crafted value
 * ({@code =HYPERLINK(...)}, {@code =cmd|...}) can exfiltrate data or run a command. Any cell whose text
 * starts with one of those characters is therefore written as an explicit string so it is stored as a
 * literal and never interpreted. Plain values (including numeric ones) are written normally so prices
 * stay numeric in the sheet.
 */
trait SpreadsheetCellWriter
{
    /**
     * The leading characters a spreadsheet treats as the start of a formula. A const (not a config
     * property) so it is a fixed security policy, not something a project can weaken via YAML.
     */
    private const FORMULA_TRIGGERS = ['=', '+', '-', '@', "\t", "\r"];

    /**
     * Write $values as one row, left to right from column A of $rowNumber.
     *
     * @param array<int, mixed> $values
     */
    protected function writeRow(Worksheet $sheet, array $values, int $rowNumber): void
    {
        $column = 1;
        foreach ($values as $value) {
            $this->writeCell($sheet, Coordinate::stringFromColumnIndex($column) . $rowNumber, $value);
            $column++;
        }
    }

    /**
     * Write one cell, neutralising a formula-trigger leading character by forcing a literal string.
     */
    protected function writeCell(Worksheet $sheet, string $coordinate, mixed $value): void
    {
        $string = is_scalar($value) ? (string) $value : '';
        if ($string !== '' && in_array($string[0], self::FORMULA_TRIGGERS, true)) {
            $sheet->setCellValueExplicit($coordinate, $string, DataType::TYPE_STRING);

            return;
        }

        $sheet->setCellValue($coordinate, $string);
    }
}
