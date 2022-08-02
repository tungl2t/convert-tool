import type { NextPage } from "next";
import Head from "next/head";
import styles from "../styles/Home.module.css";
import readXlsxFile from "read-excel-file";
import { useRef, useState } from "react";
import { Button, Textarea } from "@chakra-ui/react";

const Home: NextPage = () => {
  const inputFile = useRef<HTMLInputElement>(null);
  const [value, setValue] = useState("");
  const handleChange = (e: any) => {
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

                    if (rowData[index + 2]) {
                      report += `; ${rows[3][index + 2]}: ${
                        rowData[index + 2]
                      }`;
                    }
                    if (rowData[index + 3]) {
                      report += `; ${rows[3][index + 3]}: ${
                        rowData[index + 3]
                      }`;
                    }
                    if (rowData[index + 4]) {
                      report += `; ${rows[3][index + 4]}: ${
                        rowData[index + 4]
                      }`;
                    }
                    if (rowData[index + 5]) {
                      report += `; ${rows[3][index + 5]}: ${
                        rowData[index + 5]
                      }`;
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

  return (
    <div className={styles.container}>
      <Head>
        <title>Niteco Convert Tool</title>
        <meta name="description" content="Niteco Convert Tool" />
        <link rel="icon" href="/favicon.ico" />
      </Head>

      <main className={styles.main}>
        <Button colorScheme="teal" variant="outline" marginBottom={5} onClick={() => {inputFile?.current?.click()}}>
          Import excel file
        </Button>
        <input
          type="file"
          ref={inputFile}
          onChange={handleChange}
          onClick={(event: any) => {
            event.target.value = null;
          }}
          style={{display: 'none'}}
        />
        <Textarea value={value} isReadOnly resize="none" height={"60vh"} />
      </main>
    </div>
  );
};

export default Home;
