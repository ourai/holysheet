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
