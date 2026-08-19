import { flushPromises, mount } from "@vue/test-utils";
import { describe, expect, it, vi } from "vitest";
import type { MindooDBAppDatabaseInfo, MindooDBAppSession } from "mindoodb-app-sdk";

import { i18n } from "@/i18n";
import NewDocumentDialog from "./NewDocumentDialog.vue";

const databases: MindooDBAppDatabaseInfo[] = [
  {
    id: "teamedit",
    title: "Shared teamedit database",
    capabilities: ["read", "create", "directory", "views"],
  },
];

function createSession() {
  return {
    openDatabase: vi.fn(async () => ({
      documents: {
        listCreateKeys: vi.fn(async () => [
          { keyId: "default", isDefault: true },
          { keyId: "$publicinfos", isDefault: false },
        ]),
        getDefaultCreateKeyId: vi.fn(async () => "default"),
      },
      directory: {
        listUsers: vi.fn(async () => ["cn=Ada Lovelace/o=Acme"]),
      },
    })),
    createViewNavigator: vi.fn(async () => ({
      expandAll: vi.fn(),
      dispose: vi.fn(),
      entriesForward: vi.fn(async () => ({
        entries: [
          {
            key: "untagged",
            kind: "category",
            origin: "teamedit",
            level: 0,
            parentKey: null,
            categoryPath: [null],
            categoryValue: null,
            columnValues: {},
            position: null,
            expanded: false,
            selected: false,
            isVisible: true,
            descendantDocumentCount: 1,
            docId: null,
          },
          {
            key: "work",
            kind: "category",
            origin: "teamedit",
            level: 0,
            parentKey: null,
            categoryPath: ["Work"],
            categoryValue: "Work",
            columnValues: {},
            position: null,
            expanded: false,
            selected: false,
            isVisible: true,
            descendantDocumentCount: 1,
            docId: null,
          },
        ],
        nextPosition: null,
      })),
    })),
  };
}

function mountDialog(props: Record<string, unknown> = {}) {
  return mount(NewDocumentDialog, {
    props: {
      visible: true,
      databases,
      session: createSession() as unknown as MindooDBAppSession,
      currentUserName: "Flitz Pipe/Acme",
      currentUserCanonical: "cn=Flitz Pipe/o=Acme",
      initialType: "markdown",
      initialDatabaseId: "teamedit",
      creating: false,
      ...props,
    },
    global: {
      plugins: [i18n],
      stubs: {
        Dialog: {
          template: "<section><slot /><slot name=\"footer\" /></section>",
        },
        Button: {
          props: ["label", "disabled", "type"],
          emits: ["click"],
          template: "<button :disabled=\"disabled\" :type=\"type || 'button'\" @click=\"$emit('click')\">{{ label }}</button>",
        },
      },
    },
  });
}

describe("NewDocumentDialog", () => {
  it("emits a shared-key create draft", async () => {
    const wrapper = mountDialog();
    await flushPromises();
    expect(wrapper.text()).toContain("$publicinfos");

    await wrapper.find("input[placeholder='Document title']").setValue("Kickoff notes");
    await wrapper.find("textarea").setValue("Work\\Planning");
    const keySelect = wrapper.findAll("select").find((select) =>
      select.findAll("option").some((option) => option.element.value === "$publicinfos"),
    );
    await keySelect?.setValue("$publicinfos");

    const createButton = wrapper.findAll("button").find((button) => button.text() === "Create");
    await createButton?.trigger("click");

    expect(wrapper.emitted("create")).toEqual([
      [
        {
          databaseId: "teamedit",
          type: "markdown",
          title: "Kickoff notes",
          tags: ["Work\\Planning"],
          encryption: { mode: "shared", decryptionKeyId: "$publicinfos" },
        },
      ],
    ]);
  });

  it("lets the user encrypt for specific people and pick an existing tag", async () => {
    const wrapper = mountDialog({ initialType: "word" });
    await flushPromises();

    const peopleRadio = wrapper.find('input[type="radio"][value="people"]');
    expect(peopleRadio.exists()).toBe(true);
    await peopleRadio.setValue();

    const userSelect = wrapper.findAll("select").at(-1);
    await userSelect?.setValue("cn=Ada Lovelace/o=Acme");
    await wrapper.find("form").trigger("submit");

    await wrapper.find(".tag-tree-list__button").trigger("click");

    const createButton = wrapper.findAll("button").find((button) => button.text() === "Create");
    await createButton?.trigger("click");

    expect(wrapper.emitted("create")).toEqual([
      [
        {
          databaseId: "teamedit",
          type: "word",
          title: "",
          tags: ["Work"],
          encryption: {
            mode: "people",
            recipients: ["cn=Ada Lovelace/o=Acme"],
          },
        },
      ],
    ]);
  });

  it("does not insert null when an untagged category is present", async () => {
    const wrapper = mountDialog();
    await flushPromises();

    expect(wrapper.text()).toContain("Work");
    expect(wrapper.text()).not.toContain("Untagged");
    expect(wrapper.find("textarea").element.value).toBe("");
  });
});
