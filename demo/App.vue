<template>
  <div class="HolysheetDemo">
    <div class="HolysheetDemo-toolbar">
      <button :key="tool.key" @click.prevent="tool.handler" v-for="tool in tools">
        {{ tool.text }}
      </button>
    </div>
    <div class="HolysheetDemo-spreadsheet" id="holysheet" />
  </div>
</template>

<script lang="ts">
import { Vue, Component } from 'vue-property-decorator';

import Holysheet from 'holysheetjs';

@Component
export default class HolysheetDemo extends Vue {
  private tools: any[] = [];

  private hs: Holysheet = null as any;

  private handleOperationResult(result: any): void {
    if (!result.success) {
      alert(result.message);
    }
  }

  private created(): void {
    this.hs = new Holysheet({ column: { count: 4 }, row: { count: 4 } });

    this.tools = [
      {
        key: 'merge',
        text: '合并单元格',
        handler: () => this.handleOperationResult(this.hs.merge()),
      },
      {
        key: 'unmerge',
        text: '取消合并单元格',
        handler: () => this.handleOperationResult(this.hs.unmerge()),
      },
    ];
  }

  private mounted(): void {
    this.hs.mount('#holysheet');

    this.hs.on({
      'cell-change': cell => console.log(`cell ${cell.id} selected`, cell),
      'range-change': range => console.log(`range [${range.join(',')}] selected`, range),
      'width-change': ({ index, width }) =>
        console.log(`column ${index} width changed to ${width}`),
      'height-change': ({ index, height }) =>
        console.log(`row ${index} height changed to ${height}`),
    });

    this.hs.setSheets([{ name: 'holysheet' }]);

    this.hs.select(1, 1);
  }
}
</script>

<style>
html,
body,
.HolysheetDemo,
.HolysheetDemo-spreadsheet {
  height: 100%;
  box-sizing: border-box;
}

.HolysheetDemo {
  padding: 65px 20px 20px;
}

.HolysheetDemo-toolbar {
  margin-top: -45px;
  margin-bottom: 20px;
}

.HolysheetDemo-toolbar button {
  cursor: pointer;
}

.HolysheetDemo-toolbar button + button {
  margin-left: 10px;
}
</style>
