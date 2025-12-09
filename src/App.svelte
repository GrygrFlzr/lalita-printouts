<script lang="ts">
  import Workbook from './lib/Workbook.svelte';
  import { read } from 'xlsx';

  const acceptedFileTypes = [
    {
      file_type: 'Excel 97-2003 Workbook',
      extension: '.xls',
      mime_type: 'application/vnd.ms-excel'
    },
    {
      file_type: 'Excel Workbook',
      extension: '.xlsx',
      mime_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    },
    {
      file_type: 'Excel Macro-Enabled Workbook',
      extension: '.xlsm',
      mime_type: 'application/vnd.ms-excel.sheet.macroEnabled.12'
    },
    {
      file_type: 'Excel Binary Workbook',
      extension: '.xlsb',
      mime_type: 'application/vnd.ms-excel.sheet.binary.macroEnabled.12'
    },
    {
      file_type: 'Excel Template',
      extension: '.xltx',
      mime_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template'
    },
    {
      file_type: 'Excel Macro-Enabled Template',
      extension: '.xltm',
      mime_type: 'application/vnd.ms-excel.template.macroEnabled.12'
    },
    {
      file_type: 'Excel Add-In',
      extension: '.xlam',
      mime_type: 'application/vnd.ms-excel.addin.macroEnabled.12'
    },
    {
      file_type: 'Excel External Data Connection File',
      extension: '.xlc',
      mime_type: 'application/vnd.ms-excel'
    }
  ];
  const acceptedString = acceptedFileTypes
    .flatMap(({ extension, mime_type }) => [extension, mime_type])
    .join(',');

  let files: FileList | undefined | null = $state(null);
  $inspect(files);
</script>

<main class="container mx-auto min-h-screen bg-white px-4 py-8 not-print:shadow-md">
  <label class="flex flex-col gap-2 print:hidden">
    <span class="font-bold">Pilih file Microsoft Excel untuk diproses</span>
    <input
      bind:files
      type="file"
      accept={acceptedString}
      class="file:mr-4 file:cursor-pointer file:rounded-md file:border-b-2 file:border-b-green-900 file:bg-green-700 file:px-4 file:py-2 file:text-green-50 file:hover:bg-green-600"
    />
  </label>

  {#if files}
    {#each Array.from(files) as file}
      {#await file.arrayBuffer()}
        <p>Membaca file...</p>
      {:then binaryData}
        <Workbook workbook={read(binaryData)} />
      {:catch error}
        <h3 class="text-red-500">Gagal membaca file</h3>
        <pre><code>{error}</code></pre>
      {/await}
    {/each}
  {/if}
</main>
