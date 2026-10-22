<?php

declare(strict_types=1);

namespace SilverShop\Import\Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use RuntimeException;
use SilverShop\Import\Model\ImportProfile;
use SilverShop\Import\Source\ImportContext;
use SilverShop\Import\Source\ImportSource;
use SilverShop\Import\Source\SourceRecord;
use SilverStripe\Core\Injector\Injectable;

/**
 * Reads a spreadsheet (xlsx / csv / ods) into {@link SourceRecord}s. The header row supplies the source
 * field names; each subsequent row becomes one record. The upsert key column is the one the profile maps
 * to {@code InternalItemID} (or {@code GTIN}), falling back to the first column — so the same mapping that
 * drives the apply also identifies the key, with no extra config.
 *
 * Spreadsheets carry no per-row timestamp, so delta/idempotency relies on the engine's content hash
 * ({@code PayloadHash}), not a cursor.
 */
class ExcelImportSource implements ImportSource
{
    use Injectable;

    public function getKey(): string
    {
        return 'excel';
    }

    public function getName(): string
    {
        return 'Excel / CSV spreadsheet';
    }

    /**
     * Configured when phpspreadsheet is available (the per-profile file is checked at fetch time).
     */
    public function isConfigured(): bool
    {
        return class_exists(IOFactory::class);
    }

    public function supportsProducts(): bool
    {
        return true;
    }

    public function supportsStock(): bool
    {
        return true;
    }

    public function supportsPrices(): bool
    {
        return true;
    }

    public function fetch(ImportContext $context): iterable
    {
        $path = (string) $context->profile->SourceLocation;
        if ($path === '' || !is_file($path)) {
            throw new RuntimeException(sprintf(
                'Excel source file not found: %s',
                $path !== '' ? $path : '(no SourceLocation set on the profile)'
            ));
        }

        $rows = IOFactory::load($path)->getActiveSheet()->toArray(null, true, true, false);
        if (count($rows) < 2) {
            return;
        }

        $headers = array_map(static fn ($h): string => trim((string) $h), (array) array_shift($rows));
        $keyColumn = $this->keyColumn($context->profile, $context->domain, $headers);

        foreach ($rows as $row) {
            $fields = [];
            foreach ($headers as $i => $header) {
                if ($header === '') {
                    continue;
                }
                $fields[$header] = $row[$i] ?? null;
            }
            if ($fields === []) {
                continue;
            }

            $externalKey = $keyColumn !== '' ? trim((string) ($fields[$keyColumn] ?? '')) : '';
            if ($externalKey === '') {
                // A row with no key can't be upserted — skip it (the result workbook still shows it staged
                // only once a key is present).
                continue;
            }

            yield new SourceRecord($externalKey, $context->domain, $fields, null, $row);
        }
    }

    /**
     * The header that keys the upsert: the source field mapped to InternalItemID, then GTIN, then the first
     * column.
     *
     * @param array<int, string> $headers
     */
    private function keyColumn(ImportProfile $profile, string $domain, array $headers): string
    {
        foreach (['InternalItemID', 'GTIN'] as $target) {
            $mapping = $profile->fieldMappingsForDomain($domain)->filter('TargetField', $target)->first();
            if ($mapping && in_array($mapping->SourceField, $headers, true)) {
                return (string) $mapping->SourceField;
            }
        }

        return $headers[0] ?? '';
    }
}
