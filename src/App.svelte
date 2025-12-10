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
</script>

{#snippet errorDebug(error: unknown)}
  <fieldset class="text-md border border-green-900 px-2 pb-2">
    <legend class="text-slate-700">actual error object</legend>
    {#if error instanceof Error}
      {#if 'name' in error}
        <p><strong>name</strong>: {error.name}</p>
      {/if}
      {#if 'message' in error}
        <p><strong>message</strong>: {error.message}</p>
      {/if}
      {#if 'cause' in error}
        <p><strong>message</strong>: {error.cause}</p>
      {/if}
      <p><strong>stack</strong></p>
      {#if 'stack' in error}
        {#if typeof error.stack === 'string' && error.stack.includes('\n')}
          <ol class="list-decimal pl-8">
            {#each error.stack.split('\n').filter((line) => line.trim().length > 0) as line}
              <li class="font-mono">{line}</li>
            {/each}
          </ol>
        {:else}
          <pre class="pl-2 text-xs"><code>{error.stack}</code></pre>
        {/if}
      {/if}
    {:else}
      <pre><code>type: {typeof error}</code></pre>
      <pre><code>{error}</code></pre>
    {/if}
  </fieldset>
{/snippet}

{#snippet failed(error: unknown, reset: () => void)}
  <h3 class="text-xl text-red-500">Gagal memuat file</h3>
  <p>Harap kontak IT dengan screenshot eror bawah dan sertakan file yang bermasalah.</p>
  <button
    class="text-md cursor-pointer rounded-md border-b-2 border-b-green-900 bg-green-700 px-3 py-1.5 text-green-50 hover:bg-green-600"
    onclick={reset}>Atau klik tombol ini untuk coba lagi</button
  >
  <p>Penjelasan untuk tim IT:</p>
  <div class="flex flex-col border-l border-red-500 py-2 pl-4">
    <fieldset class="border border-black px-2 pb-2 text-sm">
      <legend class="text-slate-700">metadata</legend>
      <p><strong>getPrototypeOf</strong>: {Object.getPrototypeOf(error)}</p>
      <p><strong>properties</strong>: {Object.getOwnPropertyNames(error).join(', ')}</p>
    </fieldset>
    {@render errorDebug(error)}
  </div>
{/snippet}

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

  <svelte:boundary {failed}>
    {#if files}
      {#if files.length === 1}
        {#await files.item(0)!.arrayBuffer()}
          <p>Membaca file...</p>
        {:then binaryData}
          <Workbook workbook={read(binaryData)} {failed} />
        {:catch error}
          <h3 class="text-red-500">Gagal membaca file</h3>
          {@render errorDebug(error)}
        {/await}
      {:else}
        <p>Tolong pilih hanya satu file!</p>
      {/if}
    {/if}
  </svelte:boundary>
</main>
