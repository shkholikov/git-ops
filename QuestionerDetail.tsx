import * as React from "react";
import { ReactNode, useContext, useEffect, useRef, useState } from "react";
import { AppContext } from "../App";
import { useParams } from "react-router";
import { getAccessList, getAnswerList, getCollectionList, updateListItem } from "../utils/api";
import { IAnswer, IAnswerItem, INote, IObjectItem, IQuestion, IQuestionItem } from "../utils/IApiProps";
import {
    Accordion,
    ActionGroup,
    BarDiagram,
    Button,
    ComplexTable,
    Grid,
    H2Thin,
    Icon,
    IconLink,
    InputSelect,
    InputTextfield,
    Switch
} from "@lsg/components";
import {
    communication___news,
    interaction___checkmark,
    interaction___share,
    interaction___trash,
    object_document_document,
    symbols___error,
    symbols___infoCircle
} from "@lsg/icons";
import { getDateFromBid, hasCurrentUserPermission, setObjectStatus } from "../utils/helpers";
import jsPDF from "jspdf";
import autoTable, { CellHookData } from "jspdf-autotable";
import "jspdf-autotable";
import { useNavigate } from "react-router-dom";
import { NavigateFunction } from "react-router/dist/lib/hooks";
import Loading from "../components/banners/Loading";
import ErrorMsg from "../components/banners/ErrorMsg";
import InformationModal from "../components/dialogs/InformationModal";
import ConfirmationModal from "../components/dialogs/ConfirmationModal";
import InlineWarning from "../components/banners/InlineWarning";
import styles from "./../App.module.scss";
import { Tooltip } from "react-tooltip";

interface IQuestionerDetailProps {}

const QuestionerDetail: React.FC<IQuestionerDetailProps> = () => {
    const appContext = useContext(AppContext);
    const config = JSON.parse(localStorage.getItem("obbConfiguration"));
    const navigate: NavigateFunction = useNavigate();
    const { inspectionId } = useParams();
    const stickySaveBtn = useRef(null);
    const scrolling = useRef(document.querySelector('[data-automation-id="contentScrollRegion"]'));
    const scrollToElement = useRef(null);
    //component access matrix, Title and Name must be elements in this array!
    //btnNewInspection = StartBegehung
    const accessConfig: string[] = ["Title", "Name", "StartBegehung"];
    const [state, setState] = useState({
        errorState: false,
        errorMsg: null,
        dataLoaded: false,
        timer: null,
        inspectionNotOpened: false,
        allItemsLoaded: false,
        saveOkDialogOpen: false,
        moveNextDialogOpen: false,
        moveNextConfirmed: false,
        fillAllAnswersWarningOpen: false,
        ignoredQuestions: null,
        deleteInspectionConfirmationOpen: false,
        offlineMode: false,
        showOnlineOfflineConfirmation: false,
        triggeredQuestions: [],
        skipQuestionary: [],
        BIDyear: inspectionId.split("-")[0].substring(0, 4),
        questionDescriptionWindowOpen: false,
        openAccordionIndex: null,
        accordionMultipleOpening: true,
        questionnaireStatistics: {
            totalQuestions: 0,
            notAnsweredQuestions: 0,
            answeredQuestions: 0,
            completedPercentage: 0,
            notAnsweredChapters: []
        },
        componentPermissionMatrix: null,
        answerItem: null,
        answerCollectionItem: null,
        buildingData: null,
        currentUserGroups: appContext.initialData.currentUser.Groups,
        qdescText: "",
        qdescLink: "",
        noteText: "",
        triggerStatisticsRecalculation: false,
        showOnlyUnanswered: false,
        showOnlyUnansweredConfirmation: false,
        triggerRerender: false
    });

    useEffect(() => {
        const fetchComponentsData = async (): Promise<void> => {
            const componentPermission = await getAccessList(config.accessMatrixListName, accessConfig);
            const inspection = await getAnswerList(config.answersListName, config.queryList.answerList, `Title eq '${inspectionId}'`);
            inspection[0].ANSWERS = typeof inspection[0].ANSWERS === "string" ? JSON.parse(inspection[0].ANSWERS.trim()) : inspection[0].ANSWERS;
            inspection[0].ANSWERS = addQuestionDataToAnswerObject(appContext.initialData.questionList, inspection[0].ANSWERS);
            inspection[0].NOTES = typeof inspection[0].NOTES === "string" ? JSON.parse(inspection[0].NOTES.trim()) : inspection[0].NOTES;
            addNotesDataToAnswerObject(inspection[0]);
            const collection = appContext.initialData.collectionList.find((collection) => collection.Title === inspectionId.split("-")[1]);
            const object = appContext.initialData.objectList.find((object) => object.OID === inspectionId.split("-")[2]);
            const ignoredQuestions = await getCollectionList(
                config.collectionsListName[new Date().getFullYear()],
                config.queryList.collectionList,
                "NAME eq 'novalidation'"
            );
            if (inspection[0].STATUS === config.BegehungStatus.OPENED) {
                setState({
                    ...state,
                    componentPermissionMatrix: componentPermission,
                    ignoredQuestions: ignoredQuestions[0].DATA.split(";"),
                    triggeredQuestions: loadTriggeredQuestions(),
                    skipQuestionary: loadSkipQuestionary(),
                    answerItem: inspection[0],
                    answerCollectionItem: collection,
                    buildingData: [object],
                    triggerStatisticsRecalculation: !state.triggerStatisticsRecalculation,
                    dataLoaded: true
                });
            } else {
                setState({
                    ...state,
                    inspectionNotOpened: true,
                    dataLoaded: true
                });
            }
        };
        fetchComponentsData().catch((err: Error) => {
            console.error("An error occurred during fetching objects data in NewQuestionary component: ", err);
            setState({
                ...state,
                dataLoaded: true,
                errorState: true,
                errorMsg: err.message
            });
        });
    }, []);

    useEffect(() => {
        calculateQuestionnaireStatistics();
    }, [state.triggerStatisticsRecalculation]);

    useEffect(() => {
        if (state.dataLoaded) {
            stickySaveBtn.current = document.querySelector("#stickySaveBtn");
            stickySaveBtn.current.addEventListener("click", (): Promise<void> => void saveData(true));
        }

        //auto save
        const timer = setInterval(() => {
            if (state.offlineMode === false) {
                void saveData(false);
            }
        }, config.autosaveTime);
        setState({ ...state, timer: timer });
        //called when component unmounts to clean up the timer
        return () => clearInterval(state.timer);
    }, [state.dataLoaded]);

    function loadTriggeredQuestions() {
        const triggerQuestionString = config.triggerQuestionNegativeAnswer;
        let triggerQuestions: IQuestion[] = [];
        if (triggerQuestionString.indexOf(";") > -1) {
            const questionCollection = triggerQuestionString.split(";");
            triggerQuestions = questionCollection.map((itm: string) => cleanQuestionConditionalAnswer(itm));
        } else {
            triggerQuestions.push(cleanQuestionConditionalAnswer(triggerQuestionString));
        }
        return triggerQuestions;
    }

    function loadSkipQuestionary() {
        const skipQuestionaryString = config.skipQuestionary;
        let skippingQuestions: IQuestion[] = [];
        if (skipQuestionaryString.indexOf(",") > -1) {
            const questionCollection = skipQuestionaryString.split(",");
            questionCollection.map((itm: string) => cleanQuestionConditionalAnswer(itm));

            skippingQuestions = questionCollection;
        } else {
            skippingQuestions.push(cleanQuestionConditionalAnswer(skipQuestionaryString));
        }
        return skippingQuestions;
    }

    function cleanQuestionConditionalAnswer(itm: string): IQuestion {
        const qData = itm.split(":");
        return { Number: qData[0], Answer: qData[1] };
    }

    function addQuestionDataToAnswerObject(allQuestions: IQuestionItem[], specifiedList: string | IAnswer[]): IAnswer[] {
        const data: IAnswer[] = [];
        if (typeof specifiedList !== "string") {
            specifiedList.forEach((specifiedQuestion) => {
                for (const question of allQuestions) {
                    if (specifiedQuestion.number === question.Number) {
                        let negAnswerString = "";
                        const filterTriggeredQuestions = state.triggeredQuestions.filter(
                            (negAnswer: IQuestion) => negAnswer.Number === specifiedQuestion.number
                        );
                        negAnswerString = filterTriggeredQuestions.length > 0 ? filterTriggeredQuestions[0].Answer : question.NegativeA;
                        specifiedQuestion = {
                            number: specifiedQuestion.number,
                            answer: specifiedQuestion.answer,
                            description: question.Title,
                            possibleAnswers: question.Answer !== null && typeof question.Answer === "string" ? question.Answer.split(";") : [],
                            negativeAnswer: negAnswerString,
                            childQuestions: question.parentquestion,
                            qdesc: question.qdesc,
                            separatorNum: question.SeparatorN,
                            issueState:
                                specifiedQuestion.answer && negAnswerString
                                    ? negAnswerString.indexOf(";") > -1
                                        ? question.NegativeA.split(";").some((negAnswer) => negAnswer === specifiedQuestion.answer)
                                        : specifiedQuestion.answer === negAnswerString
                                    : false,
                            questionType: question.qtype
                        };
                        data.push(specifiedQuestion);
                        break;
                    }
                }
            });
        }
        return data;
    }

    function addNotesDataToAnswerObject(inspection: IAnswerItem) {
        if (inspection.NOTES && typeof inspection.NOTES !== "string") {
            if (inspection.NOTES.length > 0) {
                if (typeof inspection.ANSWERS !== "string") {
                    inspection.ANSWERS.forEach((itm) => {
                        itm.note = "";
                        if (typeof inspection.NOTES !== "string") {
                            inspection.NOTES.forEach((note) => {
                                if (note.number === itm.number) {
                                    itm.note = note.text;
                                }
                            });
                        }
                    });
                }
            } else {
                if (typeof inspection.ANSWERS !== "string") {
                    inspection.ANSWERS.forEach((itm) => {
                        itm.note = "";
                    });
                }
            }
        } else {
            if (typeof inspection.ANSWERS !== "string") {
                inspection.ANSWERS.forEach((itm) => {
                    itm.note = "";
                });
            }
        }
    }

    function calculateQuestionnaireStatistics() {
        if (state.answerItem) {
            const questions = state.answerItem.ANSWERS.filter(
                (item: IAnswer) => item.questionType === "question" && !item.number.startsWith(config.notIncludedInQuestionnaireInStatistics)
            );
            // let questions = state.answerItem.ANSWERS.filter((item) => item.questionType === "question");
            const chapters = state.answerItem.ANSWERS.filter((item: IAnswer) => item.questionType === "chapter");
            const notAnsweredQuestions = questions.filter((item: IAnswer) => !item.answer).length;
            const answeredQuestions = questions.filter((item: IAnswer) => !!item.answer).length;
            const notAnsweredChapters = getNotAnsweredChapters(chapters, questions);
            setState((prevState) => {
                const copyOfState = { ...prevState };
                copyOfState.questionnaireStatistics.totalQuestions = questions.length;
                copyOfState.questionnaireStatistics.notAnsweredQuestions = notAnsweredQuestions;
                copyOfState.questionnaireStatistics.answeredQuestions = answeredQuestions;
                copyOfState.questionnaireStatistics.completedPercentage = Math.round((answeredQuestions / questions.length) * 100);
                copyOfState.questionnaireStatistics.notAnsweredChapters = notAnsweredChapters;
                return copyOfState;
            });
        }
    }

    function getNotAnsweredChapters(chapters: IAnswer[], questions: IAnswer[]) {
        const notAnsweredChapters: string[] = [];
        let chapterDescription;
        questions.forEach((question) => {
            if (!question.answer) {
                chapters.forEach((chapter) => {
                    if (question.number.split(".")[0] === chapter.number.split(".")[0]) {
                        chapterDescription = chapter.number + " " + chapter.description;
                        if (!(notAnsweredChapters.indexOf(chapterDescription) > -1)) {
                            notAnsweredChapters.push(chapterDescription);
                        }
                    }
                });
            }
        });
        return notAnsweredChapters;
    }

    function createApplicationContent(): ReactNode {
        const chaptersData = state.answerItem.ANSWERS.filter((item: IAnswer) => item.questionType === "chapter");
        const data = state.showOnlyUnanswered
            ? state.answerItem.ANSWERS.sort((a: IAnswer, b: IAnswer) => a.number.localeCompare(b.number, "de", { numeric: true })).filter(
                  (item: IAnswer) => !item.answer || !item.note
              )
            : state.answerItem.ANSWERS.sort((a: IAnswer, b: IAnswer) => a.number.localeCompare(b.number, "de", { numeric: true }));
        let childQuestions;
        const pdf =
            state.answerItem.status === config.BegehungStatus.OPENED ? (
                <IconLink icon={object_document_document} onClick={() => generatePDFBlanket(data)}>
                    PDF Blanko-Checkliste
                </IconLink>
            ) : (
                ""
            );
        const deleteBtn =
            hasCurrentUserPermission(state.currentUserGroups, state.componentPermissionMatrix, "StartBegehung") === true ? (
                <Grid.Row verticalAlign="middle">
                    <Grid.Column size={2}>
                        <div style={{ textAlign: "center" }}>
                            <IconLink
                                icon={interaction___trash}
                                iconColor="error"
                                disabled={state.offlineMode === true}
                                onClick={() =>
                                    setState({
                                        ...state,
                                        deleteInspectionConfirmationOpen: true
                                    })
                                }>
                                <span style={state.offlineMode === true ? { color: "#c1c1c1" } : { color: "#ef3340" }}>Begehung Löschen</span>
                            </IconLink>
                        </div>
                    </Grid.Column>
                </Grid.Row>
            ) : (
                ""
            );
        return (
            <>
                <Grid spacing="doublesubsection">
                    <Grid.Row>
                        <Grid.Column size={8}>
                            <H2Thin>
                                {state.answerCollectionItem.NAME} für {state.buildingData[0].STRASSE + " " + state.buildingData[0].ORT}
                            </H2Thin>
                        </Grid.Column>
                        <Grid.Column size={2}>{config.exportToBlankPDF === true ? pdf : ""}</Grid.Column>
                        <Grid.Column size={2}>
                            <Switch.Group direction="vertical">
                                <Switch
                                    label="Offline Modus"
                                    value={state.offlineMode}
                                    onChange={() => setState({ ...state, showOnlineOfflineConfirmation: true })}
                                />
                                <Switch
                                    label="Unbeantwortet anzeigen"
                                    value={state.showOnlyUnanswered}
                                    onChange={() => setState({ ...state, showOnlyUnansweredConfirmation: true })}
                                />
                            </Switch.Group>
                        </Grid.Column>
                    </Grid.Row>
                    <Grid.Row>
                        <Grid.Column size={12}>{createBuildingInfo()}</Grid.Column>
                    </Grid.Row>
                    <Grid.Row verticalAlign="middle">
                        <Grid.Column size={5}>
                            <BarDiagram
                                label={"Gesamtanzahl Fragen: " + state.questionnaireStatistics.totalQuestions}
                                valueLabel={state.questionnaireStatistics.completedPercentage + "%"}
                                labelSubline={state.questionnaireStatistics.answeredQuestions + " beantwortet"}
                                valueLabelSubline={state.questionnaireStatistics.notAnsweredQuestions + " nicht beantwortet"}
                                percent={state.questionnaireStatistics.completedPercentage}
                            />
                        </Grid.Column>
                        <Grid.Column size={1}>
                            <div data-tooltip-id="diagramInfoIcon">
                                <IconLink
                                    icon={symbols___infoCircle}
                                    look="no-text"
                                    disabled={state.questionnaireStatistics.notAnsweredChapters.length === 0}
                                />
                            </div>
                            <Tooltip
                                id="diagramInfoIcon"
                                float={true}
                                style={{ fontFamily: "'Gotham', sans-serif", backgroundColor: "#EBF0F0", color: "#00333D", zIndex: 100 }}>
                                <div>
                                    <b>Kapitel mit offenen Fragen:</b>
                                </div>
                                {state.questionnaireStatistics.notAnsweredChapters.map((chapter) => {
                                    return <div>{chapter}</div>;
                                })}
                            </Tooltip>
                        </Grid.Column>
                        <Grid.Column size={6}>
                            <ActionGroup left={true}>
                                <Button
                                    disabled={state.offlineMode === true}
                                    onClick={() =>
                                        checkAnswerQuestions() === true
                                            ? setState((state) => ({ ...state, moveNextDialogOpen: true }))
                                            : setState((state) => ({ ...state, fillAllAnswersWarningOpen: true }))
                                    }>
                                    Filialcheck abschließen
                                </Button>
                                <Button look="secondary" disabled={state.offlineMode === true} onClick={() => saveData(true)}>
                                    Speichern
                                </Button>
                            </ActionGroup>
                        </Grid.Column>
                    </Grid.Row>
                    <Grid.Row>
                        <Grid.Column size={12}>
                            <Accordion.Group
                                multiple={state.accordionMultipleOpening}
                                openIndex={state.openAccordionIndex}
                                onChange={() => setState({ ...state, accordionMultipleOpening: true, openAccordionIndex: "" })}>
                                {chaptersData.map((chapter: IAnswer, index: number) => {
                                    childQuestions = data.filter((ques: IAnswer) => {
                                        if (chapter.separatorNum) {
                                            return (
                                                ques.number.split(".")[0] === chapter.number.split(".")[0] &&
                                                parseInt(ques.number.split(".")[1]) > parseInt(chapter.separatorNum.split("-")[0]) &&
                                                parseInt(ques.number.split(".")[1]) < parseInt(chapter.separatorNum.split("-")[1]) &&
                                                ques.questionType !== "chapter"
                                            );
                                        } else {
                                            return ques.number.split(".")[0] === chapter.number.split(".")[0] && ques.questionType !== "chapter";
                                        }
                                    });
                                    if (childQuestions.length > 0) {
                                        return (
                                            <Accordion key={index} title={chapter.number + " " + chapter.description}>
                                                <ComplexTable
                                                    className={styles["inspection-table"]}
                                                    columnProperties={[
                                                        { title: "Nummer", name: "number" },
                                                        { title: "Frage", name: "question" },
                                                        { title: "Antwort", name: "answer" },
                                                        { title: "Beschreibung", name: "note" }
                                                    ]}
                                                    tableBodyData={childQuestions.map((entry: IAnswer, idx: number) => {
                                                        const isChapter = entry.questionType !== null && entry.questionType === "chapter";
                                                        const RowStyle = isChapter === true ? { fontWeight: "bold", fontSize: "17px" } : {};
                                                        return {
                                                            rowId: idx.toString(),
                                                            rowData: [
                                                                <div id={entry.number} style={RowStyle}>
                                                                    {entry.number}
                                                                </div>,
                                                                !entry.qdesc ? (
                                                                    <div style={RowStyle}>{entry.description}</div>
                                                                ) : (
                                                                    <>
                                                                        <div>{entry.description}</div>
                                                                        <span data-tooltip-id={"infoIcon" + idx}>
                                                                            <IconLink
                                                                                icon={symbols___infoCircle}
                                                                                onClick={() => createQdescWindow(entry.qdesc)}>
                                                                                Ergänzende Information
                                                                            </IconLink>
                                                                        </span>
                                                                        <Tooltip
                                                                            id={"infoIcon" + idx}
                                                                            float={true}
                                                                            style={{
                                                                                fontFamily: "'Gotham', sans-serif",
                                                                                backgroundColor: "#EBF0F0",
                                                                                color: "#00333D",
                                                                                zIndex: 100
                                                                            }}>
                                                                            Ergänzende Information
                                                                        </Tooltip>
                                                                    </>
                                                                ),
                                                                isChapter === false ? (
                                                                    <InputSelect
                                                                        onChange={(selectedAnswer) =>
                                                                            answerChangedActionPerformed(entry.number, selectedAnswer)
                                                                        }
                                                                        value={entry.answer}
                                                                        options={generateOptions(entry.possibleAnswers)}
                                                                    />
                                                                ) : (
                                                                    ""
                                                                ),
                                                                isChapter === false ? (
                                                                    <InputTextfield.Stateful
                                                                        textArea={true}
                                                                        helperText={entry.issueState === true ? "Bearbeitung notwendig!!" : ""}
                                                                        defaultValue={entry.note}
                                                                        icon={entry.issueState === true ? symbols___error : ""}
                                                                        iconText="Bearbeitung notwendig!!"
                                                                        onIconClick={(event) =>
                                                                            noteChanged(entry.number, (event.target as HTMLInputElement).value)
                                                                        }
                                                                        onBlur={(event) =>
                                                                            noteChanged(entry.number, (event.target as HTMLInputElement).value)
                                                                        }
                                                                    />
                                                                ) : (
                                                                    ""
                                                                )
                                                            ]
                                                        };
                                                    })}
                                                />
                                            </Accordion>
                                        );
                                    }
                                })}
                            </Accordion.Group>
                        </Grid.Column>
                    </Grid.Row>
                </Grid>
                <Grid spacing="subsection" centeredLayout={true}>
                    {deleteBtn}
                </Grid>
            </>
        );
    }

    function checkAnswerQuestions() {
        if (config.forceFillAnswers === false) {
            return true;
        } else {
            let skipCheck = false;
            for (const skipQuestion of state.skipQuestionary) {
                state.answerItem.ANSWERS.map((itm: IAnswer) => {
                    if (skipQuestion.Number === itm.number && skipQuestion.Answer === itm.answer) {
                        skipCheck = true;
                    }
                });
            }
            if (skipCheck) {
                return true;
            } else {
                let allFilled = true;
                for (const itm of state.answerItem.ANSWERS) {
                    let negAnswerIsTriggered = false;
                    state.triggeredQuestions.every((negAnswer) => {
                        negAnswerIsTriggered = negAnswer.Number === itm.number ? itm.answer === itm.negativeAnswer : false;
                    });
                    if (
                        itm.questionType !== "chapter" &&
                        isIgnored(itm.number) === false &&
                        (itm.answer === null || itm.answer === "" || (negAnswerIsTriggered && !itm.note))
                    ) {
                        allFilled = false;
                        setState((state) => ({
                            ...state,
                            accordionMultipleOpening: false,
                            openAccordionIndex: parseInt(itm.number.split(".")[0]) - 1
                        }));
                        scrollToElement.current = document.getElementById(itm.number);
                        break;
                    }
                }
                return allFilled;
            }
        }
    }

    function isIgnored(number: string): boolean {
        if (state.ignoredQuestions.length === 0) {
            return false;
        } else {
            let ignored = false;
            for (const itm of state.ignoredQuestions) {
                if (itm === number) {
                    ignored = true;
                    break;
                }
            }
            return ignored;
        }
    }

    function generatePDFBlanket(data: IAnswer[]) {
        const objectCols = [["Begehung ID", "WE", "FILHB", "Adresse", "GS-OS Standort", "Datum"]];
        const objectData = [
            [
                inspectionId,
                state.buildingData.we,
                state.buildingData.filhb,
                state.buildingData.street + "\n" + state.buildingData.zipCode + " " + state.buildingData.city,
                state.buildingData.place,
                getDateFromBid(inspectionId)
            ]
        ];
        const cols = ["Nr.", "Frage", "", "Antwort", "Beschreibung"];
        const rows = [];
        for (const itm of data) {
            rows.push(getRowData(itm));
        }
        const doc = new jsPDF("landscape") as any;
        const totalPagesExp = "Seiten";
        doc.autoTableSetDefaults({
            headStyles: { fillColor: [255, 233, 0], textColor: [0, 65, 75] },
            styles: {
                textColor: [0, 65, 75],
                overflow: "linebreak",
                cellWidth: "wrap",
                rowPageBreak: "auto"
            },
            columnStyles: { text: { cellWidth: "linebreak" } }
        });
        doc.setTextColor(0, 65, 75);
        doc.text(state.answerCollectionItem.name, 14, 10);
        autoTable(doc, {
            head: objectCols,
            body: objectData,
            startY: 18
        });
        autoTable(doc, {
            head: [cols],
            body: rows,
            startY: doc.previousAutoTable.finalY + 10,
            columnStyles: {
                0: { cellWidth: 20 },
                1: { cellWidth: "auto" },
                2: { cellWidth: "auto" },
                3: { cellWidth: "auto" }
            },
            didDrawPage: function (data: CellHookData) {
                let str = "Seite " + doc.internal.getNumberOfPages();
                if (typeof doc.putTotalPages === "function") {
                    str = str + " von " + totalPagesExp;
                }
                doc.setFontSize(10);
                const pageSize = doc.internal.pageSize;
                const pageHeight = pageSize.height ? pageSize.height : pageSize.getHeight();
                doc.text(str, data.settings.margin.left, pageHeight - 10);
                doc.text("Begehung ID: " + inspectionId, data.settings.margin.right * 15, pageHeight - 10);
            },
            didParseCell: function (data: CellHookData) {
                if (typeof data.cell.raw === "boolean") {
                    if (data.cell.raw === true) {
                        for (const key in data.row.cells) {
                            if (Object.prototype.hasOwnProperty.call(data.row.cells, key)) {
                                data.row.cells[key].styles.fontStyle = "bold";
                            }
                        }
                    }
                    data.cell.raw = "";
                    data.cell.text = [""];
                    data.cell.width = 0;
                }
            }
        });
        if (typeof doc.putTotalPages === "function") {
            doc.putTotalPages(totalPagesExp);
        }
        doc.save("Kontrollenblankett_" + inspectionId + ".pdf");
    }

    function getRowData(entry: IAnswer) {
        const data: (string | boolean)[] = [entry.number, entry.description];
        let cross = "";
        let answers = "";
        for (const itm of entry.possibleAnswers) {
            answers += itm + "\n";
            if (entry.answer !== null && entry.answer !== "" && entry.answer === itm) {
                cross += "×\n";
            } else {
                cross += "\n";
            }
        }
        const isChapter = entry.questionType !== null && entry.questionType === "chapter";
        data.push(cross);
        data.push(answers);
        data.push(entry.note);
        data.push(isChapter);
        return data;
    }

    function createBuildingInfo(): ReactNode {
        return (
            <ComplexTable
                columnProperties={[
                    { title: "Begehung ID", name: "BID" },
                    { title: "WE", name: "we" },
                    { title: "FILHB", name: "filhb" },
                    { title: "Adresse", name: "address" },
                    { title: "GS-OS Standort", name: "place" },
                    { title: "", name: "link" }
                ]}
                tableBodyData={state.buildingData.map((entry: IObjectItem, index: string) => ({
                    rowId: index,
                    rowData: [
                        inspectionId,
                        entry.WE,
                        entry.FILHB,
                        <div>
                            {entry.STRASSE}
                            <br />
                            {entry.PLZ} {entry.ORT}
                        </div>,
                        entry.STANDORT,
                        <IconLink
                            icon={interaction___share}
                            disabled={state.offlineMode === true}
                            look="no-text"
                            onClick={(evt) => {
                                evt.stopPropagation();
                                evt.preventDefault();
                                navigate("/objects/" + entry.OID);
                            }}>
                            Weitere info
                        </IconLink>
                    ]
                }))}
            />
        );
    }

    async function saveData(showMessage: boolean): Promise<void> {
        const answerObjects = cleanUpAnswersForSave();
        await updateListItem(config.answersListName, state.answerItem.Id, {
            ANSWERS: JSON.stringify(answerObjects.answers),
            NOTES: JSON.stringify(answerObjects.notes)
        })
            .then((response) => {
                console.info("Answers have been saved. Data ID: " + response);
                if (showMessage && showMessage === true) {
                    setState((state) => ({
                        ...state,
                        saveOkDialogOpen: true,
                        triggerStatisticsRecalculation: !state.triggerStatisticsRecalculation
                    }));
                }
            })
            .catch((err) => {
                console.error("An error occurred during saving answers: " + err);
                setState({ ...state, errorState: true, errorMsg: err });
            });
    }

    async function moveNext(): Promise<void> {
        const answerObjects = cleanUpAnswersForSave();
        const answersToSolve = getAnswersToSolveAmount();
        if (answersToSolve == 0) {
            await setObjectStatus(config.buildingListName, state.buildingData.Id, config.ObjektStatus.CLOSED);
        }
        await updateListItem(config.answersListName, state.answerItem.Id, {
            ANSWERS: JSON.stringify(answerObjects.answers),
            NOTES: JSON.stringify(answerObjects.notes),
            STATUS: answersToSolve > 0 ? config.BegehungStatus.RESOLVING : config.BegehungStatus.FINISHED
        })
            .then((response) => {
                console.info("Answers have been saved and inspection has been finished. Data ID: " + response);
                navigate("/issueManagement/" + inspectionId);
            })
            .catch((err) => {
                console.error("An error occurred during saving and closing answers: " + err);
                setState({ ...state, errorState: true, errorMsg: err });
            });
    }

    function getAnswersToSolveAmount() {
        let amount = 0;
        for (const itm of state.answerItem.ANSWERS) {
            if (
                itm.answer && itm.negativeAnswer
                    ? itm.negativeAnswer.indexOf(";") > -1
                        ? itm.negativeAnswer.split(";").some((negAnswer: string) => negAnswer === itm.answer)
                        : itm.answer === itm.negativeAnswer
                    : false
            ) {
                amount++;
            }
        }
        return amount;
    }

    function cleanUpAnswersForSave() {
        const cleanedAnswers: IAnswer[] = [];
        const notes: INote[] = [];
        state.answerItem.ANSWERS.forEach((itm: IAnswer) => {
            cleanedAnswers.push({
                number: itm.number,
                answer: itm.answer
            });
            // if (itm.note && itm.note.toString() !== "") {
            itm.note = itm.note.replace(/"/g, "''");
            notes.push({
                number: itm.number ? itm.number : "",
                text: itm.note ? itm.note : "",
                dept: itm.dept ? itm.dept : "",
                subDept: itm.subDept ? itm.subDept : "",
                issueNr: itm.issueNr ? itm.issueNr : "",
                sendTo: itm.sendTo ? itm.sendTo : []
            });
            // }
        });
        return { answers: cleanedAnswers, notes: notes };
    }

    function createQdescWindow(message: string) {
        const text = message.split("link:")[0];
        const link = message.split("link:")[1];
        setState((prevState) => ({ ...prevState, questionDescriptionWindowOpen: true, qdescText: text, qdescLink: link }));
    }

    function answerChangedActionPerformed(questionNr: string, answer: string) {
        setState((prevState) => {
            prevState.answerItem.ANSWERS.forEach((itm: IAnswer) => {
                if (itm.number === questionNr) {
                    itm.issueState = itm.negativeAnswer
                        ? itm.negativeAnswer.indexOf(";") > -1
                            ? itm.negativeAnswer.split(";").some((negAnswer) => negAnswer === answer)
                            : answer === itm.negativeAnswer
                        : false;
                    if (itm.childQuestions && itm.childQuestions.toString() !== "") {
                        prefillChildAnswers(prevState.answerItem.ANSWERS, itm.childQuestions, answer);
                    }
                    return (itm.answer = answer);
                }
            });
            return { ...state, answerItem: prevState.answerItem };
        });
    }

    function prefillChildAnswers(allAnswers: IAnswer[], childQuestions: string, selectedAnswer: string) {
        const childQuestionsData = childQuestions.split("->");
        if (selectedAnswer === childQuestionsData[0]) {
            childQuestionsData[1].split(";").forEach((val) => {
                allAnswers.forEach((itm) => {
                    if (itm.number === val) {
                        return (itm.answer = childQuestionsData[2]);
                    }
                });
            });
        }
    }

    function generateOptions(list: string[]) {
        const options: { label: string; value: string }[] = [];
        list.forEach((itm) => {
            options.push({ label: itm, value: itm });
        });
        return options;
    }

    function noteChanged(questionNr: string, value: string) {
        setState((prevState) => ({
            ...prevState,
            answerItem: {
                ...prevState.answerItem,
                ANSWERS: prevState.answerItem.ANSWERS.map((item: IAnswer) => {
                    if (item.number === questionNr) {
                        return { ...item, note: value };
                    }
                    return item;
                })
            }
        }));
    }

    const inspectionNotOpenedInfo =
        state.inspectionNotOpened === true ? <InlineWarning message="Die ausgewählte Prüfung wurde bereits bearbeitet." /> : "";
    const moveNextModal = (
        <ConfirmationModal
            isOpen={state.moveNextDialogOpen}
            title="Filialcheck abschließen?"
            message="Möchten Sie den Filialcheck wirklich abschließen? Danach können Sie Ihre Antworten nicht mehr bearbeiten!"
            yesButtonText="Ja"
            noButtonText="Nein"
            onYes={() => {
                setState({ ...state, moveNextDialogOpen: false, moveNextConfirmed: true });
                void moveNext();
            }}
            onNo={() => setState({ ...state, moveNextDialogOpen: false, moveNextConfirmed: false })}
        />
    );
    const toOfflineTitle = "Offline-Modus aktivieren?";
    const toOfflineMessage =
        "Der Offline-Modus ist für die Verwendung an Orten ohne stabile Internetverbindung vorgesehen. " +
        "Im Offline-Modus wird das automatische Speichern deaktiviert. Außerdem können Sie die Begehung nicht speichern und können auch nicht zur nächsten Phase übergehen. " +
        "Im Offline-Modus klicken Sie bitte nicht auf einen anderen Menüpunkt oder aktualisieren die Seite, da sonst alle eingegebenen Daten verloren gehen.";
    const toOnlineTitle = "Wieder in online Modus wechseln?";
    const toOnlineMessage =
        "Bevor Sie wieder online gehen, stellen Sie sicher, dass Sie über eine stabile Internetverbindung verfügen. Andernfalls wird das Speichern fehlschlagen und alle Einträge sind verloren.";

    return (
        <>
            <InformationModal
                isOpen={state.saveOkDialogOpen}
                icon={<Icon icon={interaction___checkmark} color="success" size="large" />}
                title="Sicherung erfolgreich"
                message="Daten wurden erfolgreich gespeichert."
                onClose={() => setState({ ...state, saveOkDialogOpen: false })}
            />
            <InformationModal
                isOpen={state.questionDescriptionWindowOpen}
                icon={<Icon icon={communication___news} color="primary-2" size="large" />}
                title="Ergänzende Information"
                message={state.qdescText}
                link={state.qdescLink}
                onClose={() => setState((state) => ({ ...state, questionDescriptionWindowOpen: false }))}
            />
            <InformationModal
                isOpen={state.fillAllAnswersWarningOpen}
                icon={<Icon icon={symbols___error} color="note" size="large" />}
                title="Nicht alle Fragen beantwortet"
                message="Einige Fragen wurden nicht beantwortet. Bitte wählen Sie bei allen Fragen eine Antwort aus."
                onClose={() => {
                    setState((state) => ({ ...state, fillAllAnswersWarningOpen: false }));
                    if (scrollToElement) {
                        setTimeout(
                            () =>
                                scrolling.current.scrollTo({
                                    top: scrollToElement.current.getBoundingClientRect().top + scrolling.current.scrollTop,
                                    left: 0,
                                    behavior: "smooth"
                                }),
                            500
                        );
                    }
                }}
            />
            <InformationModal
                isOpen={state.deleteInspectionConfirmationOpen}
                icon={<Icon icon={symbols___error} color="note" size="large" />}
                title="Begehung löschen?"
                message="Sind Sie sicher, dass Sie die aktuelle Begehung löschen möchten? Senden Sie bitte eine Mail an strukturmanagementitflm@commerzbank.com"
                onClose={() => setState((state) => ({ ...state, deleteInspectionConfirmationOpen: false }))}
            />
            <ConfirmationModal
                isOpen={state.showOnlineOfflineConfirmation}
                title={state.offlineMode === true ? toOnlineTitle : toOfflineTitle}
                message={state.offlineMode === true ? toOnlineMessage : toOfflineMessage}
                yesButtonText="Ja"
                noButtonText="Nein"
                onYes={() => {
                    void saveData(false);
                    setState((state) => ({
                        ...state,
                        showOnlineOfflineConfirmation: false,
                        offlineMode: !state.offlineMode
                    }));
                }}
                onNo={() => setState((state) => ({ ...state, showOnlineOfflineConfirmation: false }))}
            />
            <ConfirmationModal
                isOpen={state.showOnlyUnansweredConfirmation}
                title="Unbeantwortet anzeigen"
                message={
                    state.showOnlyUnanswered
                        ? "Sind Sie damit einverstanden, dass alle Fragen angezeigt werden?"
                        : "Sind Sie damit einverstanden, dass nur unbeantwortete Fragen angezeigt werden?"
                }
                yesButtonText="Ja"
                noButtonText="Nein"
                onYes={() => setState({ ...state, showOnlyUnansweredConfirmation: false, showOnlyUnanswered: !state.showOnlyUnanswered })}
                onNo={() => setState({ ...state, showOnlyUnansweredConfirmation: false })}
            />
            {inspectionNotOpenedInfo}
            {moveNextModal}
            {state.dataLoaded === false ? (
                <Loading />
            ) : state.errorState === true ? (
                <ErrorMsg message={state.errorMsg} />
            ) : (
                createApplicationContent()
            )}
        </>
    );
};

export default QuestionerDetail;
