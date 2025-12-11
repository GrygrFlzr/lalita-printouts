<script lang="ts">
  import { onMount, tick, type Snippet } from 'svelte';
  import { utils, type WorkBook } from 'xlsx';
  interface Props {
    workbook: WorkBook;
    failed: Snippet<[error: unknown, reset: () => void]>;
  }
  const { workbook, failed }: Props = $props();
  const { SheetNames } = $derived(workbook);
  const CONFIG = {
    DEFAULT_PAGE_SIZE: 100,
    DATE_LOCALE: 'id-ID',
    TIMEZONE: 'Asia/Jakarta',
    NIK_LENGTH: 16,
    // in case someone accidentally missed a number
    TYPO_NIK_LENGTH: 15
  } as const;

  const dateFormatter = new Intl.DateTimeFormat(CONFIG.DATE_LOCALE, {
    dateStyle: 'long',
    timeZone: CONFIG.TIMEZONE
  });
  const dateKeyPattern = /^(timestamp|tanggal|tgl|date|waktu)/i;
  const maybeNIK = /\bNIK|KTP|KIA\b/;
  const formatObjectAsEntries = (maybeObject: any): [string, string | Date | number | bigint][] => {
    if (typeof maybeObject === 'object') {
      return Object.entries(maybeObject)
        .sort((a, b) => {
          const aKey = a[0].toLowerCase();
          const bKey = b[0].toLowerCase();
          if (aKey.includes('nama anak') || aKey.includes('nama pasien')) {
            return -1;
          } else if (bKey.includes('nama anak') || bKey.includes('nama pasien')) {
            return 1;
          }
          return 0;
        })
        .map(([key, value]) => {
          switch (typeof value) {
            case 'number':
            case 'bigint':
              if (maybeNIK.test(key)) {
                const valueAsString = value.toString();
                const valueLength = valueAsString.length;
                if (valueLength === CONFIG.NIK_LENGTH || valueLength === CONFIG.TYPO_NIK_LENGTH) {
                  return [key, valueAsString];
                }
              }
              // fallback to just returning the bigint/number
              return [key, value];
            case 'boolean':
              // autocast to indonesian
              return [key, value ? 'Ya' : 'Tidak'];
            case 'string': {
              if (dateKeyPattern.test(key)) {
                // UTC Dates that SheetJS doesn't convert?
                const maybeDate = new Date(value);
                const isValidDate = !isNaN(maybeDate.getTime());
                if (isValidDate) {
                  return [key, maybeDate];
                }
              }
              return [key, value];
            }
            case 'object':
              if (value instanceof Date && !isNaN(value.getTime())) {
                // rely on SheetJS's Date parsing
                return [key, value];
              }
              return [key, JSON.stringify(value, null, 2)];
            // really strange cases that should never happen, ideally
            case 'function':
            case 'symbol':
              return [key, value.toString()];
            case 'undefined':
              return [key, 'undefined'];
            default:
              return [key, `${value}`];
          }
        });
    }
    return [];
  };

  const isDateEntry = (item: [string, any]): item is [string, Date] => item[1] instanceof Date;
  const findTimestamp = (item: [string, Date]) => item[0].toLowerCase().includes('timestamp');
  const trySortByDate = (maybeSortable: [string, string | Date | number | bigint][][]) => {
    if (maybeSortable.length === 0) {
      // base case: no need to sort empty array
      return maybeSortable;
    }
    // otherwise, try sort (don't mutate original)
    return [...maybeSortable].sort((a, b) => {
      const aDates = a.filter(isDateEntry);
      const bDates = b.filter(isDateEntry);
      if (aDates.length > 0 && bDates.length > 0) {
        const aEntry = aDates.find(findTimestamp);
        const bEntry = bDates.find(findTimestamp);
        if (typeof aEntry !== 'undefined' && typeof bEntry !== 'undefined') {
          // both have timestamps
          const aDate = aEntry[1];
          const bDate = bEntry[1];
          return bDate.getTime() - aDate.getTime();
        } else if (typeof aEntry !== 'undefined' && typeof bEntry === 'undefined') {
          // only a has timestamp, move it to earlier in order
          return -1;
        } else if (typeof aEntry === 'undefined' && typeof bEntry !== 'undefined') {
          // only b has timestamp, move it to earlier in order
          return 1;
        }
      }
      // otherwise, sort descending by number of keys
      return b.length - a.length;
    });
  };

  const chunk = <T,>(list: T[], chunkSize: number) => {
    if (chunkSize <= 0) {
      throw new Error('chunkSize must be positive');
    }
    const pageCount = Math.ceil(list.length / chunkSize);
    return [...Array(pageCount)].map((_, index) => {
      const start = index * chunkSize;
      const end = start + chunkSize;
      return list.slice(start, end);
    });
  };

  let isPrinting = $state(false);
  let printIndex = $state(0);
  const resetPrint = () => {
    isPrinting = false;
    printIndex = 0;
  };
  const triggerPrint = (itemIndex: number) => {
    isPrinting = true;
    printIndex = itemIndex;
    (async () => {
      await tick();
      window.print();
    })();
  };

  // svelte-ignore state_referenced_locally
  let selectedSheetName: undefined | string = $state(
    SheetNames.length > 0 ? SheetNames[0] : undefined
  );
  let searchQuery: string = $state('');
  const lowercaseSearchQuery = $derived(searchQuery.toLowerCase());
  let currentPage = $state(0);
  let pageSize = $state(CONFIG.DEFAULT_PAGE_SIZE);
  const selectedSheet = $derived(
    typeof selectedSheetName === 'undefined'
      ? null
      : utils.sheet_to_json(workbook.Sheets[selectedSheetName])
  );
  onMount(() => {
    window.addEventListener('afterprint', resetPrint);
    return () => {
      window.removeEventListener('afterprint', resetPrint);
    };
  });
</script>

<div class="print:hidden">
  <p>Berhasil membaca {SheetNames.length} lembar sheet dalam file.</p>
  <label class="flex max-w-80 flex-col">
    {#if SheetNames.length > 1}
      <span class="font-bold">Pilih salah satu sheet:</span>
    {/if}
    <select
      class="rounded-sm border border-green-500 bg-green-100 px-4 py-2 text-green-950"
      bind:value={selectedSheetName}
      onchange={() => {
        searchQuery = '';
        currentPage = 0;
      }}
    >
      {#each workbook.SheetNames as sheetName (sheetName)}
        <option value={sheetName}>{sheetName}</option>
      {/each}
    </select>
  </label>
</div>

<svelte:boundary {failed}>
  {#if selectedSheetName}
    {#key selectedSheetName}
      {#if selectedSheet && selectedSheet.length > 0}
        <label class="flex max-w-80 flex-col print:hidden">
          <span class="font-bold">Cari data:</span>
          <input
            type="text"
            class="rounded-sm border border-green-500 bg-green-100 px-4 py-2 text-green-950"
            bind:value={searchQuery}
            oninput={() => {
              currentPage = 0;
            }}
          />
        </label>
        {@const originalList = selectedSheet.map(formatObjectAsEntries)}
        {@const filteredList =
          searchQuery === ''
            ? originalList
            : originalList.filter((item) =>
                item.some(([_key, value]) => {
                  const searchableValue =
                    value instanceof Date
                      ? dateFormatter.format(value)
                      : typeof value === 'string'
                        ? value
                        : String(value);
                  return searchableValue.toLowerCase().includes(lowercaseSearchQuery);
                })
              )}
        {@const sortedList = trySortByDate(filteredList)}
        {@const chunkedList = chunk(sortedList, pageSize)}
        {@const pageEntries = chunkedList[currentPage]}
        <fieldset class="print:hidden">
          <legend>
            Halaman
            <strong>{currentPage + 1}</strong>
            dari
            <strong>{chunkedList.length}</strong>
            {searchQuery.length > 0 ? '(difilter)' : ''}
          </legend>
          {#if chunkedList.length > 1}
            <button
              aria-label="Halaman sebelumnya"
              disabled={currentPage === 0}
              onclick={() => {
                currentPage--;
              }}
              class="rounded-md border-b border-b-green-900 bg-green-700 px-3 py-1.5 text-xs text-green-50 not-disabled:cursor-pointer not-disabled:hover:bg-green-600 disabled:border-b-gray-600 disabled:bg-gray-300 disabled:text-gray-500"
              >&lt; Halaman Sebelumnya</button
            >
            <button
              aria-label="Halaman selanjutnya"
              disabled={currentPage === chunkedList.length - 1}
              onclick={() => {
                currentPage++;
              }}
              class="rounded-md border-b border-b-green-900 bg-green-700 px-3 py-1.5 text-xs text-green-50 not-disabled:cursor-pointer not-disabled:hover:bg-green-600 disabled:border-b-gray-600 disabled:bg-gray-300 disabled:text-gray-500"
              >Halaman Selanjutnya &gt;</button
            >
          {/if}
        </fieldset>
        <div class="mt-6 grid grid-cols-1 gap-6 not-print:sm:grid-cols-2">
          {#each pageEntries as entry, index}
            <div
              class:hidden={isPrinting ? index !== printIndex : false}
              class="relative grid grid-cols-1 gap-y-2 rounded-sm p-4 not-print:bg-green-50 not-print:shadow-sm print:grid-cols-2 print:gap-x-6 print:gap-y-1"
            >
              <div class="absolute top-0 right-0 p-2 print:hidden">
                <button
                  aria-label="Print entri ini"
                  onclick={() => triggerPrint(index)}
                  class="text-md cursor-pointer rounded-md border-b-2 border-b-green-900 bg-green-700 px-3 py-1.5 text-green-50 hover:bg-green-600"
                  >Print</button
                >
              </div>
              {#each entry as [key, value] (key)}
                {@const lowerCaseKey = key.toLowerCase()}
                <div
                  class="flex break-inside-avoid flex-col divide-y divide-green-500 print:divide-black"
                >
                  {#if lowerCaseKey.includes('timestamp') && value instanceof Date}
                    <span class="text-sm print:text-xs">Tanggal masuk data</span>
                    <span class="ml-2">{dateFormatter.format(value)}</span>
                  {:else if lowerCaseKey.includes('nama anak') || lowerCaseKey.includes('nama pasien')}
                    <span class="text-sm font-semibold print:text-xs">{key}</span>
                    <span class="ml-2 font-bold">{value}</span>
                  {:else if typeof value === 'string' && value.includes('\n')}
                    <span class="text-sm print:text-xs">{key}</span>
                    <div class="ml-2">
                      {#each value.split('\n') as line}
                        {#if line.trim().length > 0}
                          <p>{line}</p>
                        {:else}
                          <br />
                        {/if}
                      {/each}
                    </div>
                  {:else if (typeof value === 'number' || typeof value === 'bigint') && key.includes('(')}
                    {@const parenIndex = key.indexOf('(')}
                    {@const unitlessKey = key.slice(0, parenIndex).trim()}
                    {@const unit = key.slice(parenIndex + 1).replace(/\)$/, '')}
                    <span class="text-sm print:text-xs">{unitlessKey}</span>
                    <p class="ml-2">{value} {unit}</p>
                  {:else}
                    <span class="text-sm print:text-xs">{key}</span>
                    <span class="ml-2">
                      {#if value instanceof Date}
                        {dateFormatter.format(value)}
                      {:else}
                        {value}
                      {/if}
                    </span>
                  {/if}
                </div>
              {/each}
            </div>
          {:else}
            <p>Tidak ada hasil pencarian.</p>
          {/each}
        </div>
      {:else}
        <p>Pilihan sheet anda <strong>{selectedSheetName}</strong> tidak memiliki data</p>
      {/if}
    {/key}
  {:else}
    <p>Pilih file dengan paling sedikit satu sheet di atas.</p>
  {/if}
</svelte:boundary>
