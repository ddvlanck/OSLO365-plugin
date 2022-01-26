<template>
  <vl-content-footer id="content-footer">
    <vl-layout>
      <vl-grid class="vl-grid--v-center">
        <vl-column width="6">
          <vl-button mod-block @click="insertNote('footnote')">Voetnoot</vl-button>
        </vl-column>
        <vl-column width="6">
          <vl-button mod-block @click="insertNote('endnote')">Eindnoot</vl-button>
        </vl-column>
      </vl-grid>
    </vl-layout>
  </vl-content-footer>
</template>

<script lang="ts">
import Vue from "vue";
import { onInsertNoteClicked } from "../../../utils/Utils";
import EventBus from "../../../utils/EventBus";
import { IOsloItem } from "src/oslo/IOsloItem";

export default Vue.extend({
  data: () => {
    return {
      radioTile: {} as IOsloItem
    };
  },
  methods: {
    async insertNote(which: string) {
      if (Object.keys(this.radioTile).length > 0) {
        await onInsertNoteClicked(this.radioTile, which);
      }
    }
  },
  mounted() {
    EventBus.$on("onRadioTileChanged", (data: IOsloItem) => {
      this.radioTile = data;
    });
  }
});
</script>

<style lang="scss">
#content-footer {
  z-index: 1;
  position: fixed;
  bottom: 0;
  width: 100%;
  text-align: center;
}

#content-footer div {
  background: #ffe615 !important;
}
</style>
