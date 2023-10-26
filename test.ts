const handleScroll = () => {
      const scrollY = scrolling.scrollTop;
      if (scrollY >= 160 && !showStickySaveBtn) {
        setShowStickySaveBtn(true);
      } else if (scrollY < 160 && showStickySaveBtn) {
        setShowStickySaveBtn(false);
      }
    };

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
                onYes={getOnlyUnanswered}
                onNo={() => setState({ ...state, showOnlyUnansweredConfirmation: false })}
            />



                      function getOnlyUnanswered() {
        if (!state.showOnlyUnanswered) {
            //deep copy
            let currentAnswerItem = JSON.parse(JSON.stringify(state.answerItem));
            currentAnswerItem.ANSWERS = currentAnswerItem.ANSWERS.filter((item: IAnswer) => !item.answer && item.questionType === "question");
            setState((state) => ({
                ...state,
                showOnlyUnanswered: !state.showOnlyUnanswered,
                answerItemWithUnanswered: currentAnswerItem,
                showOnlyUnansweredConfirmation: false
            }));
        } else {
            setState((state) => ({
                ...state,
                showOnlyUnanswered: !state.showOnlyUnanswered,
                showOnlyUnansweredConfirmation: false
            }));
        }
        void saveData(false);
    }
