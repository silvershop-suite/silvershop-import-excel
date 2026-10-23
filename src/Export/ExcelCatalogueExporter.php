<?php

declare(strict_types=1);

namespace SilverShop\Import\Excel\Export;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use SilverShop\Model\Variation\Variation;
use SilverShop\Page\Product;
use SilverShop\Page\ProductCategory;
use SilverStripe\Core\Config\Configurable;
use SilverStripe\Core\Injector\Injectable;
use SilverStripe\ORM\DataObject;

/**
 * Full catalogue export in the grouped import shape: one {@code record_type=product} row per product,
 * followed by a {@code record_type=variant} row per variation, so the client can correct the sheet and
 * re-import it through the same {@link \SilverShop\Import\Excel\ExcelImportSource} (export → correct →
 * re-apply). Variant rows repeat the parent's SKU ({@code InternalItemID}) and carry their own GTIN,
 * price and option axes; the parent-product row carries title, base price, category path and brand.
 *
 * Columns are configurable but default to the full roundtrip schema. Only database fields that exist on
 * the record are written (so GTIN appears only when product-identity is installed), and every cell is
 * written formula-injection-safe via {@see SpreadsheetCellWriter}.
 */
class ExcelCatalogueExporter
{
    use Injectable;
    use Configurable;
    use SpreadsheetCellWriter;

    /**
     * Must match the import reader's defaults so a straight export re-imports cleanly.
     */
    private const PATH_SEPARATOR = ' > ';
    private const PATHS_SEPARATOR = '|';
    private const MAX_AXES = 3;

    /**
     * Columns to export, in order. Overridable in YAML.
     *
     * @var array<int, string>
     */
    private static array $columns = [
        'record_type', 'InternalItemID', 'Title', 'BasePrice', 'GTIN', 'Price', 'Quantity',
        'Category', 'Brand',
        'Option1Name', 'Option1Value', 'Option2Name', 'Option2Value', 'Option3Name', 'Option3Value',
    ];

    /**
     * Export the whole product catalogue (products + their variations) to $path in the grouped shape.
     * Returns the number of rows written. The $domain argument is accepted for signature compatibility
     * with the generic exporter contract; the catalogue export is always product-centric.
     */
    public function export(string $path, string $domain = 'Product'): int
    {
        $columns = array_values(array_filter(
            (array) static::config()->get('columns'),
            static fn ($c): bool => is_string($c) && $c !== ''
        ));

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $this->writeRow($sheet, $columns, 1);

        $rowNumber = 2;
        $count = 0;
        foreach (Product::get() as $product) {
            $variations = $this->variationsOf($product);
            $this->writeRow(
                $sheet,
                $this->rowValues($columns, $this->productRow($product, $variations !== [])),
                $rowNumber++
            );
            $count++;

            foreach ($variations as $variation) {
                $this->writeRow(
                    $sheet,
                    $this->rowValues($columns, $this->variationRow($product, $variation)),
                    $rowNumber++
                );
                $count++;
            }
        }

        (new Xlsx($spreadsheet))->save($path);

        return $count;
    }

    /**
     * Order an associative row by the configured column list, filling absent columns with ''.
     *
     * @param array<int, string> $columns
     * @param array<string, string> $row
     * @return array<int, string>
     */
    private function rowValues(array $columns, array $row): array
    {
        return array_map(static fn (string $column): string => $row[$column] ?? '', $columns);
    }

    /**
     * @return array<string, string>
     */
    private function productRow(Product $product, bool $hasVariations): array
    {
        return [
            'record_type' => 'product',
            'InternalItemID' => (string) $product->getField('InternalItemID'),
            'Title' => (string) $product->getField('Title'),
            'BasePrice' => (string) $product->getField('BasePrice'),
            'GTIN' => $product->hasField('GTIN') ? (string) $product->getField('GTIN') : '',
            // Stock lives on the sellable unit: a simple product carries it; a varianted product's stock
            // is on its variants, so the parent row leaves Quantity blank.
            'Quantity' => $hasVariations ? '' : $this->stockLevel($product),
            'Category' => $this->categoryPaths($product),
            'Brand' => $this->brandTitle($product),
        ];
    }

    /**
     * @return array<string, string>
     */
    private function variationRow(Product $product, Variation $variation): array
    {
        return [
            'record_type' => 'variant',
            // Variant rows repeat the parent SKU so the reader can resolve the owning product.
            'InternalItemID' => (string) $product->getField('InternalItemID'),
            'GTIN' => $variation->hasField('GTIN') ? (string) $variation->getField('GTIN') : '',
            'Price' => (string) $variation->getField('Price'),
            'Quantity' => $this->stockLevel($variation),
        ] + $this->variationAxes($product, $variation);
    }

    /**
     * The product's variations, or [] when the variations extension isn't applied.
     *
     * @return array<int, Variation>
     */
    private function variationsOf(Product $product): array
    {
        if (!$product->hasMethod('Variations')) {
            return [];
        }

        /** @var array<int, Variation> $variations */
        $variations = $product->Variations()->toArray();

        return $variations;
    }

    /**
     * The buyable's current stock level as a string, or '' when the stock module isn't installed or the
     * level is unlimited/unset. getStockLevel() comes from silvershop/stock (a suggest-only dependency),
     * so it is resolved dynamically and guarded — the exporter never hard-depends on the stock module.
     */
    private function stockLevel(DataObject $record): string
    {
        if (!$record->hasMethod('getStockLevel')) {
            return '';
        }

        $getStockLevel = 'getStockLevel';
        $level = $record->$getStockLevel();

        return $level === null ? '' : (string) $level;
    }

    /**
     * Pipe-separated full category paths (e.g. {@code Men > Shoes|Clearance}) for every category the
     * product belongs to, mirroring the import reader's {@code Category} column.
     */
    private function categoryPaths(Product $product): string
    {
        if (!$product->hasMethod('ProductCategories')) {
            return '';
        }

        $paths = [];
        foreach ($product->ProductCategories() as $category) {
            $segments = [];
            $node = $category;
            // Walk up the tree while the ancestor is still a product category.
            while ($node instanceof ProductCategory && $node->exists()) {
                array_unshift($segments, (string) $node->getField('Title'));
                $node = $node->Parent();
            }
            if ($segments !== []) {
                $paths[] = implode(self::PATH_SEPARATOR, $segments);
            }
        }

        return implode(self::PATHS_SEPARATOR, $paths);
    }

    private function brandTitle(Product $product): string
    {
        if (!$product->hasField('BrandID') || (int) $product->getField('BrandID') === 0) {
            return '';
        }

        $brand = $product->hasMethod('Brand') ? $product->Brand() : null;

        return $brand instanceof DataObject && $brand->exists() ? (string) $brand->getField('Title') : '';
    }

    /**
     * Option{N}Name / Option{N}Value columns for a variation, ordered by the product's declared
     * attribute types so the same axis lands in the same column across every variant.
     *
     * @return array<string, string>
     */
    private function variationAxes(Product $product, Variation $variation): array
    {
        if (!$variation->hasMethod('AttributeValues')) {
            return [];
        }

        $axes = [];
        $slot = 1;
        $types = $product->hasMethod('VariationAttributeTypes') ? $product->VariationAttributeTypes() : null;

        if ($types !== null && $types->exists()) {
            foreach ($types as $type) {
                if ($slot > self::MAX_AXES) {
                    break;
                }
                $value = $variation->AttributeValues()->filter('TypeID', $type->ID)->first();
                $axes["Option{$slot}Name"] = (string) ($type->getField('Label') ?: $type->getField('Name'));
                $axes["Option{$slot}Value"] = $value ? (string) $value->getField('Value') : '';
                $slot++;
            }

            return $axes;
        }

        foreach ($variation->AttributeValues() as $value) {
            if ($slot > self::MAX_AXES) {
                break;
            }
            $type = $value->Type();
            $axes["Option{$slot}Name"] = $type && $type->exists()
                ? (string) ($type->getField('Label') ?: $type->getField('Name'))
                : '';
            $axes["Option{$slot}Value"] = (string) $value->getField('Value');
            $slot++;
        }

        return $axes;
    }
}
