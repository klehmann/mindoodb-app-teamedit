import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";

import DocumentRevisionDialog from "./DocumentRevisionDialog.vue";

const entries = [
  {
    revisionId: "rev-2",
    timestamp: Date.UTC(2026, 3, 29, 12, 0),
    heads: ["head-2"],
    publicKey: "public-key-2",
    publicKeyFingerprint: "fingerprint-2",
    isDeleted: false,
    isCurrent: true,
    summary: "Updated body",
  },
  {
    revisionId: "rev-1",
    timestamp: Date.UTC(2026, 3, 28, 12, 0),
    heads: ["head-1"],
    publicKey: "public-key-1",
    publicKeyFingerprint: "fingerprint-1",
    isDeleted: false,
    isCurrent: false,
    summary: "Created",
  },
];

function mountDialog(props = {}) {
  return mount(DocumentRevisionDialog, {
    props: {
      visible: true,
      entries,
      loading: false,
      errorMessage: null,
      currentRevisionId: "rev-2",
      ...props,
    },
    global: {
      stubs: {
        Dialog: {
          template: "<section><slot /><slot name=\"footer\" /></section>",
        },
        Button: {
          props: ["label", "disabled"],
          emits: ["click"],
          template: "<button :disabled=\"disabled\" @click=\"$emit('click')\">{{ label }}</button>",
        },
      },
    },
  });
}

describe("DocumentRevisionDialog", () => {
  it("emits the selected revision id when opening a row", async () => {
    const wrapper = mountDialog();
    const rows = wrapper.findAll(".revision-row");

    await rows[1].trigger("click");
    const buttons = wrapper.findAll("button");
    await buttons[buttons.length - 1].trigger("click");

    expect(wrapper.emitted("select")).toEqual([["rev-1"]]);
  });

  it("emits on row double click", async () => {
    const wrapper = mountDialog();

    await wrapper.findAll(".revision-row")[1].trigger("dblclick");

    expect(wrapper.emitted("select")).toEqual([["rev-1"]]);
  });

  it("shows an empty state", () => {
    const wrapper = mountDialog({
      entries: [],
      currentRevisionId: null,
    });

    expect(wrapper.text()).toContain("No revisions are available");
  });
});
