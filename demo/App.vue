<template>
  <div id="holysheet" class="HolysheetDemo"></div>
</template>

<script lang="ts">
import { Vue, Component } from 'vue-property-decorator';

import Holysheet from 'holysheetjs';

@Component
export default class HolysheetDemo extends Vue {
  private hs: Holysheet = null as any;

  private created(): void {
    this.hs = new Holysheet({ column: { count: 10 }, row: { count: 20 } });
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

    this.hs.select(5, 5);
  }
}
</script>

<style>
html,
body,
.HolysheetDemo {
  height: 100%;
}
</style>
