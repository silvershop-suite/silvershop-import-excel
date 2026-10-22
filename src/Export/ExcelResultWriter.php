<?php

declare(strict_types=1);

namespace SilverShop\Import\Excel\Export;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use SilverShop\Import\Model\ImportProfile;
use SilverShop\Import\Model\ImportStagingRecord;
use SilverStripe\Core\Injector\Injectable;

/**
 * The correction half of the roundtrip: writes a profile+domain's staged rows back to a workbook with
 * {@code _Status} and {@code _Error} columns appended, so the client fixes the flagged rows in place and
 * re-uploads. Re-apply is SKU-keyed and idempotent (unchanged rows no-op; fixed rows retry).
 */
class ExcelResultWriter
{
    use Injectable;

    /**
     * Write the staging rows for a profile + domain to $path. Returns the number of data rows written.
     */
    public function write(ImportProfile $profile, string $domain, string $path): int
    {
        $fieldKeys = [];
        $data = [];
        foreach (ImportStagingRecord::get()->filter(['ProfileID' => $profile->ID, 'Domain' => $domain]) as $row) {
            /** @var ImportStagingRecord $row */
            $fields = $row->fieldsArray();
            foreach (array_keys($fields) as $key) {
                $fieldKeys[(string) $key] = true;
            }
            $data[] = ['fields' => $fields, 'status' => (string) $row->Status, 'error' => (string) $row->Error];
        }

        $headers = array_keys($fieldKeys);
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray(array_merge($headers, ['_Status', '_Error']), null, 'A1');

        $rowNumber = 2;
        foreach ($data as $entry) {
            $line = [];
            foreach ($headers as $header) {
                $value = $entry['fields'][$header] ?? '';
                $line[] = is_scalar($value) ? $value : '';
            }
            $line[] = $entry['status'];
            $line[] = $entry['error'];
            $sheet->fromArray($line, null, 'A' . $rowNumber);
            $rowNumber++;
        }

        (new Xlsx($spreadsheet))->save($path);

        return count($data);
    }
}
