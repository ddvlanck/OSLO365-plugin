<template>
  <vl-content-footer id="content-footer">
    <vl-layout>
      <vl-grid v-vl-align:center mod-stacked>
        <vl-column width="6">
          <vl-button mod-block @click="insertNote('footnote')">Voetnoot</vl-button>
        </vl-column>
        <vl-column width="6">
          <vl-button mod-block @click="insertNote('endnote')">Eindnoot</vl-button>
        </vl-column>
        <vl-column>
          <a>Toevoegen aan woordenboek</a>
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
    async insertNote(which) {
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
@import "../css/style.scss";

#content-footer {
  z-index: 1;
  position: fixed;
  bottom: 0;
}

#content-footer div {
  background: #ffe615 !important;
}

#content-footer a {
  cursor: grab;
}
</style>
