/**
 * @file GuidHelper.ts
 * Helper methods to generate unique id (Guid)
 *
 * @copyright 2017 Olivier Carpentier
 * Released under MIT licence
 */
export default class GuidHelper {
  /**
   * @function
   * Generates a GUID
   */
  public static getGuid(): string {
    return this.s4() + this.s4() + '-' + this.s4() + '-' + this.s4() + '-' +
      this.s4() + '-' + this.s4() + this.s4() + this.s4();
  }

  /**
   * @function
   * Generates a GUID part
   */
  private static s4(): string {
      return Math.floor((1 + Math.random()) * 0x10000)
        .toString(16)
        .substring(1);
  }
}