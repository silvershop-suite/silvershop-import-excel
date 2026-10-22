<?php

declare(strict_types=1);

namespace SilverShop\Import\Excel\Export;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use SilverShop\Model\Variation\Variation;
use SilverShop\Page\Product;
use SilverStripe\Core\Config\Configurable;
use SilverStripe\Core\Injector\Injectable;
use SilverStripe\ORM\DataObject;

/**
 * Full catalogue export: dumps buyables to a workbook as an editing template / pre-change backup. The
 * columns are configurable; only database fields that exist on the record are written (so GTIN appears
 * only when product-identity is installed).
 */
class ExcelCatalogueExporter
{
    use Injectable;
    use Configurable;

    /**
     * Columns to export, in order. Overridable in YAML.
     *
     * @var array<int, string>
     */
    private static array $columns = ['InternalItemID', 'Title', 'BasePrice'];

    /**
     * Export all buyables of a domain to $path. Returns the number of rows written.
     */
    public function export(string $path, string $domain = 'Product'): int
    {
        /** @var class-string<DataObject> $class */
        $class = $domain === 'Variation' ? Variation::class : Product::class;
        $columns = array_values(array_filter(
            (array) static::config()->get('columns'),
            static fn ($c): bool => is_string($c) && $c !== ''
        ));

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray($columns, null, 'A1');

        $rowNumber = 2;
        $count = 0;
        foreach ($class::get() as $record) {
            /** @var DataObject $record */
            $line = [];
            foreach ($columns as $column) {
                $line[] = $record->hasField($column) ? (string) $record->getField($column) : '';
            }
            $sheet->fromArray($line, null, 'A' . $rowNumber);
            $rowNumber++;
            $count++;
        }

        (new Xlsx($spreadsheet))->save($path);

        return $count;
    }
}
