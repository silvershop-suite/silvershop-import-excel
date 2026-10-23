<?php

declare(strict_types=1);

namespace SilverShop\Import\Excel\Tests;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use RuntimeException;
use SilverShop\Import\Excel\ExcelImportSource;
use SilverShop\Import\Excel\Export\ExcelCatalogueExporter;
use SilverShop\Import\Model\ImportProfile;
use SilverShop\Import\Source\ImportContext;
use SilverShop\Import\Source\SourceRecord;
use SilverShop\Model\Variation\AttributeType;
use SilverShop\Model\Variation\AttributeValue;
use SilverShop\Model\Variation\Variation;
use SilverShop\Page\Product;
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

    /**
     * Build a throwaway profile pointing at $path and return the records the reader yields for $domain.
     *
     * @return array<int, SourceRecord>
     */
    private function fetchFrom(string $path, string $domain): array
    {
        $profile = ImportProfile::create();
        $profile->SourceKey = 'excel';
        $profile->SourceLocation = $path;
        $profile->write();

        return iterator_to_array((new ExcelImportSource())->fetch(new ImportContext($profile, $domain)), false);
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

    public function testRecordTypeColumnRoutesRowsByDomain(): void
    {
        $path = $this->writeWorkbook([
            ['record_type', 'SKU', 'GTIN', 'Option1Name', 'Option1Value'],
            ['product', 'ABC', '', '', ''],
            ['variant', 'ABC', '4006381333931', 'Colour', 'Red'],
            ['variant', 'ABC', '4012345678901', 'Colour', 'Blue'],
        ]);

        $products = $this->fetchFrom($path, 'Product');
        $variants = $this->fetchFrom($path, 'Variation');
        unlink($path);

        // Product domain takes only the product row; Variation domain only the variant rows.
        $this->assertCount(1, $products);
        $this->assertSame('ABC', $products[0]->get('SKU'));
        $this->assertCount(2, $variants);
        $this->assertSame('4006381333931', $variants[0]->get('GTIN'));
        $this->assertSame('Red', $variants[0]->get('Option1Value'));
    }

    public function testFlatSheetYieldsProductsButNoVariations(): void
    {
        // A sheet with no record_type column is a flat product-only catalogue.
        $path = $this->writeWorkbook([
            ['SKU', 'Name'],
            ['ABC', 'Widget'],
        ]);
        $products = $this->fetchFrom($path, 'Product');
        $variants = $this->fetchFrom($path, 'Variation');
        unlink($path);

        $this->assertCount(1, $products);
        $this->assertCount(0, $variants);
    }

    public function testCatalogueExporterRoundTripsGroupedShape(): void
    {
        $product = Product::create();
        $product->InternalItemID = 'RT-1';
        $product->Title = 'Roundtrip';
        $product->BasePrice = 10.00;
        $product->write();

        $hasVariations = $product->hasMethod('Variations') && class_exists(Variation::class);
        if ($hasVariations) {
            $type = AttributeType::create();
            $type->Name = 'Colour';
            $type->Label = 'Colour';
            $type->write();
            $product->VariationAttributeTypes()->add($type);

            $value = AttributeValue::create();
            $value->TypeID = $type->ID;
            $value->Value = 'Red';
            $value->write();

            $variation = Variation::create();
            $variation->ProductID = $product->ID;
            $variation->InternalItemID = 'RT-1-RED';
            $variation->Price = 12.00;
            $variation->write();
            $variation->AttributeValues()->add($value);
        }

        $path = (string) tempnam(sys_get_temp_dir(), 'exp') . '.xlsx';
        ExcelCatalogueExporter::create()->export($path);

        $products = $this->fetchFrom($path, 'Product');
        $variants = $this->fetchFrom($path, 'Variation');
        unlink($path);

        // The product row round-trips with its grouped shape.
        $this->assertCount(1, $products);
        $this->assertSame('product', $products[0]->get('record_type'));
        $this->assertSame('RT-1', $products[0]->get('InternalItemID'));

        if ($hasVariations) {
            $this->assertCount(1, $variants);
            $this->assertSame('variant', $variants[0]->get('record_type'));
            // Variant row repeats the parent SKU and carries the option axis.
            $this->assertSame('RT-1', $variants[0]->get('InternalItemID'));
            $this->assertSame('Colour', $variants[0]->get('Option1Name'));
            $this->assertSame('Red', $variants[0]->get('Option1Value'));
        }
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
