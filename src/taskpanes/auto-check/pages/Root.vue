<template>
  <vl-layout v-if="!scanned">
    <vl-grid mod-stacked>
      <vl-column>
        <vl-introduction>
          Voer een documentscan uit om te kijken welke woorden uit de OSLO Knowledge Graph herkent worden in je document
        </vl-introduction>
      </vl-column>
      <vl-column>
        <vl-action-group mod-align-center mod-collapse-s>
          <vl-button mod-wide @click="scan">Start scan</vl-button>
        </vl-action-group>
      </vl-column>
    </vl-grid>
  </vl-layout>
  <div v-else>
    <vl-layout>
      <vl-grid mod-stacked v-if="scanned && results.length > 0">
        <vl-column>
          <vl-title tag-name="h2">
            Gevonden definities voor <span class="vl-u-mark">{{ shownWord.text }}</span>
          </vl-title>
        </vl-column>
        <vl-column>
          <vl-action-group mod-space-between>
            <vl-button mod-icon-before icon="nav-left-light" @click="previous" :mod-disabled="resultIndex === 0"
              >Vorige</vl-button
            >
            <vl-introduction>{{ resultIndex + 1 }} / {{ results.length }}</vl-introduction>
            <vl-button
              mod-icon-after
              icon="nav-right-light"
              @click="next"
              :mod-disabled="resultIndex === results.length - 1"
              >Volgende</vl-button
            >
          </vl-action-group>
        </vl-column>
        <vl-column id="ResultBox">
          <search-result-card
            v-for="(hit, index) of shownWordDefinitions"
            :key="`${hit.reference}-${index}`"
            :value="hit"
            :id="`radio-tile-${index}`"
            :title="hit.label"
            :description="hit.description"
            :url="hit.reference"
          />
        </vl-column>
      </vl-grid>
      <vl-grid mod-stacked v-if="scanned && results.length === 0">
        <vl-column>
          <vl-introduction>Er werden geen overeenkomsten gevonden in OSLO voor het document.</vl-introduction>
        </vl-column>
        <vl-column v-vl-align:center>
          <vl-button @click="scan">Opnieuw scannen</vl-button>
        </vl-column>
      </vl-grid>
    </vl-layout>
    <content-footer v-if="scanned && results.length > 0" />
  </div>
</template>

<script lang="ts">
import Vue from "vue";
import { searchDocument, getDefinitions, selectWordInDocument } from "../auto-check";
import searchResultCard from "../../../general-components/search-result-card/search-result-card.vue";
import contentFooter from "../components/content-footer-auto-check-pane.vue";
import { IOsloItem } from "src/oslo/IOsloItem";

export default Vue.extend({
  name: "root",
  components: { searchResultCard, contentFooter },
  data: () => {
    return {
      scanned: false,
      searching: false,
      resultIndex: 0,
      results: [] as Word.Range[],
      shownWord: {} as Word.Range,
      shownWordDefinitions: [] as IOsloItem[],
      selectedDefinition: {} as IOsloItem
    };
  },
  methods: {
    async scan() {
      this.searching = true;
      this.scanned = true;

      this.results = await searchDocument();
      this.shownWord = this.results[this.resultIndex];
      this.shownWordDefinitions = getDefinitions(this.shownWord);

      this.searching = false;
    },
    next() {
      if (this.resultIndex + 1 <= this.results.length - 1) {
        this.resultIndex++;
        this.updateDisplayedWord();
      }
    },
    previous() {
      if (this.resultIndex - 1 >= 0) {
        this.resultIndex--;
        this.updateDisplayedWord();
      }
    },
    updateDisplayedWord() {
      this.shownWord = this.results[this.resultIndex];
      this.shownWordDefinitions = getDefinitions(this.shownWord);
    },
    selectShownWordInDocument() {
      selectWordInDocument(this.shownWord);
    }
  },
  watch: {
    shownWord(newValue) {
      selectWordInDocument(newValue);
    }
  },
});
</script>

<style lang="scss">
@import "../css/style.scss";

body {
  overflow-x: hidden;
}

#ResultBox {
  margin-bottom: 120px;
}

/* width */
::-webkit-scrollbar {
  width: 10px;
} /* Track */
::-webkit-scrollbar-track {
  background: lightgrey;
  border-radius: 10px;
} /* Handle */
::-webkit-scrollbar-thumb {
  background: grey;
  border-radius: 10px;
}
</style>
