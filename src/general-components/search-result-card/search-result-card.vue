<template>
  <vl-radio-tile
    class="radio-tile"
    name="oslo-search-result-radio"
    v-model="selectedRadioTile"
    :value="value"
    :id="id"
    :title="title"
    :info="description"
  >
    <a :href="url">{{ url }}</a>
  </vl-radio-tile>
</template>

<script lang="ts">
import Vue from "vue";
import EventBus from "../../utils/EventBus";
export default Vue.extend({
  props: {
    value: {
      type: Object,
      default: () => {}
    },
    id: {
      type: String,
      default: "radio-tile"
    },
    title: {
      type: String,
      default: ""
    },
    description: {
      type: String,
      default: ""
    },
    url: {
      type: String,
      default: ""
    }
  },
  data: () => {
    return {
      selectedRadioTile: {}
    };
  },
  watch: {
    selectedRadioTile(update) {
      EventBus.$emit("onRadioTileChanged", update);
    }
  }
});
</script>

<style lang="scss">
@import "./style.scss";

.radio-tile a {
  word-wrap: break-word;
}

/** Prevents jumping to top when clicking a radio tile */
input[type="radio"] {
  visibility: hidden;
}
</style>
