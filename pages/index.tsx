import type { NextPage } from "next";
import Head from "next/head";
import styles from "../styles/Home.module.css";
import readXlsxFile from "read-excel-file";
import { useRef, useState } from "react";
import { Button, Flex, Textarea } from "@chakra-ui/react";
import { Cell } from "read-excel-file/types";

const Home: NextPage = () => {
  const inputFileEl = useRef<HTMLInputElement>(null);
  const textAreaEl = useRef<HTMLTextAreaElement>(null);
  const [value, setValue] = useState("");
  const [label, setLabel] = useState("Copy");

  const handleStatusData = (
    status: Cell,
    statusValue: string | number | boolean | DateConstructor
  ) => {
    return statusValue === "x" ? `; ${status}` : `; ${status}: ${statusValue}`;
  };

  const handleChange = (e: any) => {
    setLabel("Copy");
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      readXlsxFile(file).then((rows) => {
        let report = "Project summary report: \n\n";
        if (rows[0][1] && rows[0][3]) {
          report += `${rows[0][1]}: ${rows[0][3]} \n\n`;
        }
        if (rows[1][1] && rows[1][3]) {
          report += `${rows[1][1]}: ${rows[1][3]} \n\n`;
        }

        if (rows[2]) {
          const listProjects = rows[2];

          listProjects.forEach((project, index) => {
            if (index % 6 === 1) {
              report += `\n${project}: \n\n`;

              for (let rowIndex = 4; rowIndex < rows.length; rowIndex++) {
                const rowData = rows[rowIndex];
                if (rowData[index]) {
                  if (
                    rowData[index + 2] ||
                    rowData[index + 3] ||
                    rowData[index + 4] ||
                    rowData[index + 5]
                  ) {
                    report += `${rowData[index]}`;
                    if (rowData[index + 1]) {
                      report += `: ${rowData[index + 1]}`;
                    }

                    /**
                     * Status Index:
                     *  2: Done
                     *  3: In-progress
                     *  4: Customer Review
                     *  5: Pending input
                     */
                    for (let statusIndex = 2; statusIndex <= 5; statusIndex++) {
                      if (rowData[index + statusIndex]) {
                        report += handleStatusData(
                          rows[3][index + statusIndex],
                          rowData[index + statusIndex]
                        );
                      }
                    }
                    report += "\n";
                  }
                }
              }
            }
          });
        }

        setValue(report);
      });
    }
  };

  const onCopyText = () => {
    const textData = textAreaEl?.current?.value ?? "";
    if (textData) {
      textAreaEl?.current?.select();
      navigator.clipboard.writeText(textData);
      setLabel("Copied!");
    }
  };

  return (
    <div className={styles.container}>
      <Head>
        <title>Niteco Convert Tool</title>
        <meta name="description" content="Niteco Convert Tool" />
        <link rel="icon" href="/favicon.ico" />
      </Head>

      <main className={styles.main}>
        <Flex
          maxWidth={600}
          w={"95%"}
          flexDirection={"column"}
          justifyContent={"center"}
          alignContent={"center"}
        >
          <Button
            colorScheme="teal"
            variant="outline"
            marginBottom={5}
            onClick={() => {
              inputFileEl?.current?.click();
            }}
          >
            Import Excel File
          </Button>
          <input
            type="file"
            ref={inputFileEl}
            onChange={handleChange}
            onClick={(event: any) => {
              event.target.value = null;
            }}
            style={{ display: "none" }}
          />
          <Textarea
            ref={textAreaEl}
            value={value}
            isReadOnly
            resize="none"
            height={"60vh"}
            w={"100%"}
            variant="outline"
          />
          <Button
            marginTop={5}
            colorScheme="teal"
            variant="outline"
            onClick={onCopyText}
          >
            {label}
          </Button>
        </Flex>
      </main>
    </div>
  );
};

export default Home;
