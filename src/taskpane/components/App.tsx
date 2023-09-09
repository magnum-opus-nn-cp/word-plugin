import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import { FC, PropsWithChildren, useCallback, useEffect, useRef } from "react";
import { ScoreType, useCreateProcess, useProcess, useText } from "../api/process";
import { useSingleTimeout } from "../hooks/useSingleTimeout";

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global Word, require */

const processFunction = async (
  context: Word.RequestContext,
  search_text: string,
  comment_text: string,
  score: number
) => {
  search_text.replace("\u0005", "");
  try {
    const range = context.document.body
      .search(search_text.slice(0, 255).split("\\").join(""), {
        ignorePunct: true,
        ignoreSpace: true,
      })
      .getFirstOrNullObject();

    await context.sync();

    if (!range.isNullObject) {
      range.insertComment(comment_text + "\n" + "Точность: " + score.toFixed(2));
    }
  } catch (e) {
    //
  }

  // return context.document.body
  //   .search(search_text.slice(0, 255).split("\\").join(""), {
  //     ignorePunct: true,
  //     ignoreSpace: true,
  //   })
  //   .getFirst()
  //   .insertComment(comment_text + "\n" + "Точность: " + score.toFixed(2));
};

const typeKey: ScoreType = "bert";

export const getEntriesFromDescription = (description: string = "") => {
  const matches = Array.from(description.matchAll(/<span[^>]+>(.*?)<\/span>/gi));

  const array = matches
    .map((match) => {
      const entry = match[0];
      const text = match[1];
      const value = Number(Array.from(entry.matchAll(/data-value=\\?"?([.\d]+)\\?"?/gi))[0]?.[1]) || 0;

      return {
        text,
        value,
      };
    })
    .filter((i) => i.value >= 0.01);

  return [...Array.from(new Map(array.map((item) => [item["text"], item])).values())];
};

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export const PROCESS_POLLING_MS = 2000;

export const App: FC<PropsWithChildren<AppProps>> = () => {
  const {
    data: createProcessResponse,
    mutateAsync: createProcess,
    isLoading: createProcessLoading,
  } = useCreateProcess();
  const processId = createProcessResponse?.id;

  const {
    data: process,
    isFetching: processFetching,
    refetch: refetchProcess,
  } = useProcess({
    processId: processId || "",
    config: {
      enabled: !!processId,
    },
  });

  const timeout = useSingleTimeout();

  useEffect(() => {
    if (processId) {
      const startPolling = () => {
        timeout.set(async () => {
          const { data: process } = await refetchProcess();
          if (process && process.current < process.total) {
            startPolling();
          }
        }, PROCESS_POLLING_MS);
      };

      if (processId) {
        startPolling();
      }
    }
  }, [processId, refetchProcess, timeout]);

  // const textId = process?.current === process?.total ? process?.texts[0].id : undefined;
  const textId = 3184;

  const {
    data: textEntity,
    isFetching: textFetching,
    refetch: refetchText,
  } = useText({
    textId: textId || 0,
    type: typeKey,
    config: {
      enabled: !!textId,
    },
  });

  const answer = textEntity?.score?.[typeKey]?.answer;
  const metric = textEntity?.score?.[typeKey]?.metric;
  const summary = textEntity?.summary;

  const initRef = useRef(false);

  useEffect(() => {
    if (!textEntity || initRef.current) {
      return;
    }

    initRef.current = true;

    Word.run(async (context) => {
      const entries = getEntriesFromDescription((textEntity?.description?.[typeKey]?.text as string) || "");
      for (const entry of entries) {
        await processFunction(context, entry.text, "", entry.value);
      }
    });
  }, [textEntity]);

  const isLoading =
    createProcessLoading ||
    !!processId ||
    processFetching ||
    !!(process && process.current < process.total) ||
    textFetching;

  const onClick = useCallback(() => {
    Word.run(function (context) {
      // Insert your code here. For example:
      const documentBody = context.document.body;
      context.load(documentBody);
      return context.sync().then(async () => {
        await refetchText();

        // const text = documentBody.text;
        // await createProcess({
        //   text,
        // });
      });
    });
  }, []);

  return (
    <div className="ms-welcome">
      <Header message="Анализ текстовых пресс-релизов" />

      <main className="ms-welcome__main">
        <p>
          Позволяет оценить кредитный рейтинг компании на основе пресс-релиза с выделением в тексте ключевых фраз с
          использованием различных методов.
        </p>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={onClick}>
          Отправить
        </DefaultButton>

        {isLoading && !textEntity && <p>Обработка...</p>}

        {textEntity && (
          <>
            <h3>Результат оценки</h3>
            <p>
              Оценочный рейтинг: {answer} <br />
              {metric && `Точность: ${metric.toFixed(2)}`}
            </p>
            <h4>Краткое содержание</h4>
            <p>{summary}</p>
          </>
        )}
      </main>
    </div>
  );
};
