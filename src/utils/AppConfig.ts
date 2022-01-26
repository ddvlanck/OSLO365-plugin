export namespace AppConfig {
  /** The URL of the Oslo data file. Retrieved with a simple GET. This static resource is replaced with a dynamic backend endpoint in the production environment. */
  export const dataFileUrl = "/data/oslo_terminology.json";

  /** Set true to enable some trace messages to help debugging. */
  export const trace = true;
}
