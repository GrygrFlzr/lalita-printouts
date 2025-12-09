<script lang="ts">
  import { onMount, tick } from 'svelte';
  import { utils, type WorkBook } from 'xlsx';
  interface Props {
    workbook: WorkBook;
  }
  const { workbook }: Props = $props();
  const { SheetNames } = $derived(workbook);

  const excelEpochMs = Date.UTC(1900, 0, 1);
  const dateFormatter = new Intl.DateTimeFormat('id-ID', {
    dateStyle: 'long',
    timeZone: 'Asia/Jakarta'
  });
  const parseExcelDate = (daysSinceEpoch: number) =>
    new Date(excelEpochMs + daysSinceEpoch * 24 * 60 * 60 * 1000);
  const formatObject = (maybeObject: any): Record<string, any> => {
    if (typeof maybeObject === 'object') {
      let importantField = null;
      for (const [key, value] of Object.entries(maybeObject)) {
        const lowerCaseKey = key.toLowerCase();
        if (
          importantField === null &&
          (lowerCaseKey.includes('nama anak') || lowerCaseKey.includes('nama pasien'))
        ) {
          // only use first appearance
          importantField = key;
        }
        if (typeof value === 'number') {
          if (lowerCaseKey.includes('timestamp') || lowerCaseKey.includes('tanggal')) {
            maybeObject[key] = parseExcelDate(value);
          }
        } else if (typeof value === 'boolean') {
          maybeObject[key] === value ? 'Ya' : 'Tidak';
        }
      }
      if (importantField) {
        return {
          [importantField]: maybeObject[importantField],
          ...maybeObject
        };
      }
      return maybeObject;
    }
    return {};
  };

  const trySortByDate = (maybeSortable: Record<string, any>[]) => {
    // TODO: Handle case where some rows have timestamp and others don't
    if (maybeSortable.length === 0) {
      return maybeSortable;
    }
    const firstSample = maybeSortable[0];
    const timestampKey = Object.keys(firstSample).find((name) =>
      name.toLowerCase().includes('timestamp')
    );
    if (timestampKey) {
      maybeSortable.sort((a, b) => {
        const aDate: Date = a[timestampKey];
        const bDate: Date | undefined = b[timestampKey];
        if (typeof bDate === 'undefined') {
          // push undated rows to back of list
          return -1;
        }
        return bDate.getTime() - aDate.getTime();
      });
    }
    return maybeSortable;
  };

  const chunk = (list: Record<string, any>[], chunkSize: number) => {
    const pageCount = Math.ceil(list.length / chunkSize);
    return [...Array(pageCount)].map((_, index) => {
      const start = index * chunkSize;
      const end = start + chunkSize;
      return list.slice(start, end);
    });
  };

  let isPrinting = $state(false);
  let printIndex = $state(0);
  const triggerPrint = (itemIndex: number) => {
    isPrinting = true;
    printIndex = itemIndex;
    (async () => {
      await tick();
      window.print();
    })();
  };

  // svelte-ignore state_referenced_locally
  let selectedSheetName: undefined | string = $state();
  let searchQuery: string = $state('');
  let currentPage = $state(0);
  let pageSize = $state(100);
  const selectedSheet = $derived(
    typeof selectedSheetName === 'undefined'
      ? null
      : utils.sheet_to_json(workbook.Sheets[selectedSheetName])
  );
  onMount(() => {
    const resetPrint = () => {
      isPrinting = false;
      printIndex = 0;
    };
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
      }}
    >
      {#each workbook.SheetNames as sheetName, index (sheetName)}
        <option value={sheetName}>{sheetName}</option>
      {/each}
    </select>
  </label>
</div>

{#if selectedSheetName}
  {#key selectedSheetName}
    {#if selectedSheet && selectedSheet.length > 0}
      <label class="flex max-w-80 flex-col print:hidden">
        <span class="font-bold">Cari data:</span>
        <input
          type="text"
          bind:value={searchQuery}
          class="rounded-sm border border-green-500 bg-green-100 px-4 py-2 text-green-950"
          onchange={() => {
            currentPage = 0;
          }}
        />
      </label>
      {@const originalList = selectedSheet.map(formatObject)}
      {@const filteredList =
        searchQuery === ''
          ? originalList
          : originalList.filter((item) =>
              Object.values(item).some(
                // at least one value contains the substring
                (item) =>
                  typeof item === 'string' && item.toLowerCase().includes(searchQuery.toLowerCase())
              )
            )}
      {@const sortedList = trySortByDate(filteredList)}
      {@const chunkedList = chunk(sortedList, pageSize)}
      {@const pageEntries = chunkedList[currentPage]}
      <fieldset class="print:hidden">
        <legend>
          Halaman <strong>{currentPage + 1}</strong> dari <strong>{chunkedList.length}</strong>
        </legend>
        {#if chunkedList.length > 1}
          <button
            disabled={currentPage === 0}
            onclick={() => {
              currentPage--;
            }}
            class="rounded-md border-b border-b-green-900 bg-green-700 px-3 py-1.5 text-xs text-green-50 not-disabled:cursor-pointer not-disabled:hover:bg-green-600 disabled:border-b-gray-600 disabled:bg-gray-300 disabled:text-gray-500"
            >&lt; Halaman Sebelumnya</button
          >
          <button
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
                onclick={() => triggerPrint(index)}
                class="text-md cursor-pointer rounded-md border-b-2 border-b-green-900 bg-green-700 px-3 py-1.5 text-green-50 hover:bg-green-600"
                >Print</button
              >
            </div>
            {#each Object.entries(entry) as [key, value] (key)}
              {@const lowerCaseKey = key.toLowerCase()}
              <div
                class="flex break-inside-avoid flex-col divide-y divide-green-500 print:divide-black"
              >
                {#if lowerCaseKey.includes('timestamp')}
                  <span class="text-sm print:text-xs">Tanggal masuk data</span>
                  <span class="ml-2">{dateFormatter.format(value)}</span>
                {:else if lowerCaseKey.includes('nama anak') || lowerCaseKey.includes('nama pasien')}
                  <span class="text-sm font-semibold print:text-xs">{key}</span>
                  <span class="ml-2 font-bold">{value}</span>
                {:else if typeof value === 'string' && value.includes('\n')}
                  <span class="text-sm print:text-xs">{key}</span>
                  <p class="ml-2">{value}</p>
                {:else if typeof value === 'number' && key.includes('(')}
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
                    {:else if typeof value === 'boolean'}
                      {value ? 'Ya' : 'Tidak'}
                    {:else}
                      {value}
                    {/if}
                  </span>
                {/if}
              </div>
            {/each}
          </div>
        {/each}
      </div>
    {:else}
      <p>Pilihan sheet anda <strong>{selectedSheetName}</strong> tidak memiliki data</p>
    {/if}
  {/key}
{:else}
  <p>Pilih file dengan paling sedikit satu sheet di atas.</p>
{/if}
