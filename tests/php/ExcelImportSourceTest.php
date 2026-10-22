<?php

declare(strict_types=1);

namespace SilverShop\Import\Excel\Tests;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use RuntimeException;
use SilverShop\Import\Excel\ExcelImportSource;
use SilverShop\Import\Model\ImportProfile;
use SilverShop\Import\Source\ImportContext;
use SilverShop\Import\Source\SourceRecord;
use SilverStripe\Dev\SapphireTest;

class ExcelImportSourceTest extends SapphireTest
{
    protected $usesDatabase = true;

    private function writeWorkbook(array $rows): string
    {
        $path = (string) tempnam(sys_get_temp_dir(), 'imp') . '.xlsx';
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->fromArray($rows, null, 'A1');
        (new Xlsx($spreadsheet))->save($path);

        return $path;
    }

    public function testReadsRowsAsSourceRecords(): void
    {
        $path = $this->writeWorkbook([
            ['SKU', 'Name', 'Price'],
            ['ABC', 'Widget', '9.99'],
            ['DEF', 'Gadget', '19.00'],
        ]);

        $profile = ImportProfile::create();
        $profile->SourceKey = 'excel';
        $profile->SourceLocation = $path;
        $profile->write();

        $records = iterator_to_array((new ExcelImportSource())->fetch(new ImportContext($profile, 'Product')), false);
        unlink($path);

        $this->assertCount(2, $records);
        $this->assertInstanceOf(SourceRecord::class, $records[0]);
        // Key column falls back to the first column (SKU) when nothing is mapped yet.
        $this->assertSame('ABC', $records[0]->externalKey);
        $this->assertSame('Widget', $records[0]->get('Name'));
        $this->assertSame('DEF', $records[1]->externalKey);
    }

    public function testMissingSourceFileThrows(): void
    {
        $profile = ImportProfile::create();
        $profile->SourceKey = 'excel';
        $profile->SourceLocation = '';
        $profile->write();

        $this->expectException(RuntimeException::class);
        iterator_to_array((new ExcelImportSource())->fetch(new ImportContext($profile, 'Product')), false);
    }
}
