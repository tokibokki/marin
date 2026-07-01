const {
  Document, Packer, Paragraph, TextRun,
  HeadingLevel, AlignmentType, ShadingType, BorderStyle
} = require('docx');
const fs = require('fs');

// ─── Helpers ────────────────────────────────────────────────────────────────

function h1(text, color = '1F4E79') {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, color, bold: true, font: 'Arial' })],
    spacing: { before: 360, after: 120 }
  });
}

function h2(text, color = '1A5276') {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, color, bold: true, font: 'Arial' })],
    spacing: { before: 280, after: 80 }
  });
}

function quoteBlock(text, bg = 'D6EAF8', textColor = '1A5276') {
  return new Paragraph({
    children: [new TextRun({ text: `"${text}"`, bold: true, italics: true, size: 26, color: textColor, font: 'Arial' })],
    shading: { fill: bg, type: ShadingType.CLEAR },
    indent: { left: 480, right: 480 },
    spacing: { before: 160, after: 160 },
    border: { left: { style: BorderStyle.THICK, size: 12, color: textColor } }
  });
}

function row(labelText, bodyText, labelColor = '1A5276') {
  return new Paragraph({
    children: [
      new TextRun({ text: `${labelText}:  `, bold: true, color: labelColor, font: 'Arial', size: 22 }),
      new TextRun({ text: bodyText, font: 'Arial', size: 22, color: '2C2C2C' })
    ],
    spacing: { before: 60, after: 60 },
    indent: { left: 240 }
  });
}

function essayBox(text, bg = 'E8F8F5', borderColor = '1E8449') {
  return new Paragraph({
    children: [
      new TextRun({ text: 'ESSAY SENTENCE:  ', bold: true, color: borderColor, font: 'Arial', size: 22 }),
      new TextRun({ text, italics: true, font: 'Arial', size: 22, color: '1E5631' })
    ],
    shading: { fill: bg, type: ShadingType.CLEAR },
    indent: { left: 480, right: 480 },
    spacing: { before: 120, after: 200 },
    border: { left: { style: BorderStyle.THICK, size: 12, color: borderColor } }
  });
}

function divider(color = 'AAAAAA') {
  return new Paragraph({
    children: [new TextRun('')],
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color } },
    spacing: { before: 120, after: 120 }
  });
}

function gap() {
  return new Paragraph({ children: [new TextRun('')], spacing: { before: 60, after: 60 } });
}

// ─── Content: 1984 Quotes ───────────────────────────────────────────────────

const quotes1984 = [
  {
    n: 1,
    quote: 'Big Brother is watching you.',
    speaker: 'Phrase on Party posters throughout Oceania; the regime\'s public declaration of omniscience, first encountered in the novel\'s opening pages.',
    meaning: 'A direct address to every citizen that the state is perpetually monitoring their actions, movements, and thoughts at all times.',
    techniques: 'Second-person address ("you") collapses distance between propaganda and the individual, making the threat feel personal and inescapable. Personification of "Big Brother" anthropomorphises abstract state power into a patriarchal figure, suggesting intimate, familial surveillance. Euphemism: "watching" implies care (as a parent watches a child) when it actually denotes control and threat. Repetition throughout the text enacts through form the omnipresence it describes.',
    significance: 'Establishes that the Party\'s power is primarily psychological. The Party need not watch every citizen at every moment — citizens simply need to believe they are watched. This makes self-policing inevitable, anticipating Foucault\'s theory of the panopticon, where visibility itself becomes a mechanism of control rather than a product of it.',
    context: 'Drawn from Orwell\'s study of Stalinist propaganda; Stalin\'s image similarly appeared on posters throughout the USSR, transforming political structure into a human symbol of authority. Written in 1948 at the dawn of the Cold War, Orwell feared this model of state surveillance would spread to the West.',
    essay: 'Through the omnipresent slogan "Big Brother is watching you," Orwell constructs a regime whose power operates not through physical force but through the psychological certainty of observation, compelling citizens to internalise the gaze of the state and police themselves.'
  },
  {
    n: 2,
    quote: 'WAR IS PEACE. FREEDOM IS SLAVERY. IGNORANCE IS STRENGTH.',
    speaker: 'The three official slogans of the Party, inscribed on the Ministry of Truth\'s facade and recurring throughout the novel.',
    meaning: 'Three oxymoronic declarations that directly contradict logical meaning, forcing citizens to hold contradictory ideas simultaneously.',
    techniques: 'Paradox/oxymoron — each slogan is logically impossible, enacting doublethink by form. Tricolon — the three-part structure gives rhythmic, rhetorical power. Capitalisation presents them as absolute truths, beyond question. The aphoristic structure makes them memorable and feel self-evident despite being nonsensical.',
    significance: 'Embodies "doublethink" — the Party\'s capacity to make citizens accept logical contradictions. By controlling language at its most basic level, the Party demonstrates that truth itself is a political construction. The slogans are also self-fulfilling: perpetual war does produce social stability; enslaving citizens to the collective does remove the burden of individual responsibility.',
    context: 'Directly echoes Nazi propaganda ("Arbeit Macht Frei" — Work Makes You Free — appeared above concentration camp gates) and Stalinist doublespeak, where purges were called "liberations." Orwell, who had witnessed propaganda at work in the Spanish Civil War, understood that language is the first casualty of authoritarian power.',
    essay: 'Orwell\'s deployment of paradox in the Party slogans "WAR IS PEACE. FREEDOM IS SLAVERY. IGNORANCE IS STRENGTH." exemplifies his central argument that totalitarian regimes do not merely suppress truth — they reconstruct it, training citizens to hold contradictory beliefs simultaneously through the mechanism of doublethink.'
  },
  {
    n: 3,
    quote: 'If you want a picture of the future, imagine a boot stamping on a human face — forever.',
    speaker: 'O\'Brien, Part 3, during Winston\'s interrogation in the Ministry of Love. This is O\'Brien\'s explicit statement of what the Party\'s vision of the future is.',
    meaning: 'The Party\'s goal is not prosperity, progress, or ideology — it is the exercise of pure, perpetual power over individuals. The future is simply domination, without end.',
    techniques: 'Extended metaphor — boot and face become universal symbols of oppressor and oppressed. "Forever" is devastating in its finality, closing off any possibility of change or escape. Present tense ("imagine") implicates the reader directly, making the image visceral and immediate. The anonymity of "a human face" — not named, not individual — makes the victim universal.',
    significance: 'Strips away any pretence that the Party pursues power for an ideological end. O\'Brien is explicit that the goal is power for its own sake. This distinguishes Orwell\'s dystopia from all others — unlike Huxley\'s Brave New World, where control is maintained through pleasure, Orwell\'s regime uses pain and humiliation because it enjoys them. There is no arc toward improvement.',
    context: 'Echoes Stalinist show trials, where confessions were extracted purely to demonstrate the state\'s power over the individual psyche. The "boot" imagery deliberately echoes the Nazi jackboot — a symbol of fascist authority. Writing post-WWII, Orwell was responding to the question of whether political systems could ever become "finished" in their domination.',
    essay: 'O\'Brien\'s declaration that the future is "a boot stamping on a human face — forever" strips the Party of any ideological pretence, revealing to audiences that totalitarian power is not a means to an end but an end in itself — a vision of perpetual domination that forecloses all hope of resistance or redemption.'
  },
  {
    n: 4,
    quote: 'Freedom is the freedom to say that two plus two make four. If that is granted, all else follows.',
    speaker: 'Winston, writing in his diary, Part 1. This is his articulation of what freedom actually means at its most fundamental level.',
    meaning: 'True freedom is the ability to assert objective, verifiable fact. From this basic epistemological freedom — the right to trust one\'s own perception — all other freedoms follow logically.',
    techniques: 'Repetition of "freedom" gives the sentence philosophical, definitional weight. The mathematical example (2+2=4) is deliberately simple — if even this basic truth can be denied, nothing is safe. The conditional "if that is granted" implies it has not been granted. "All else follows" positions rational truth as the root from which all other liberties grow, making its loss catastrophic.',
    significance: 'Winston understands that the Party\'s deepest attack is epistemological — on the right to know reality. This anticipates the tragedy of the novel\'s conclusion: by the close, Winston will accept that 2+2=5. The quote shows Winston grasps the stakes intellectually long before he loses the battle emotionally.',
    context: 'Reflects Orwell\'s belief in the primacy of objective truth, developed through his experience reporting the Spanish Civil War, where newspapers on all sides simply fabricated events. He wrote that the most dangerous thing totalitarianism does is make people doubt the evidence of their own senses. The simple maths example democratises the argument — this is not an elite concern but a universal one.',
    essay: 'Winston\'s diary entry that "freedom is the freedom to say that two plus two make four" crystallises Orwell\'s thesis that the most fundamental form of oppression is not physical but epistemological — the destruction of objective truth — positioning the individual\'s right to basic rational thought as the bedrock on which all other freedoms depend.'
  },
  {
    n: 5,
    quote: 'We shall meet in the place where there is no darkness.',
    speaker: 'O\'Brien, in Winston\'s recurring dream; he says this with prophetic certainty. The cruel irony is later revealed: "the place where there is no darkness" is the perpetually lit interrogation rooms of the Ministry of Love.',
    meaning: 'Initially appears to be a promise of a haven — a place of light and liberation. On revelation, it means the opposite: the Ministry of Love\'s bright, never-switched-off fluorescent lights are where Winston will be destroyed.',
    techniques: 'Dramatic irony — the reader eventually understands the true meaning; the revelation is devastating. Foreshadowing — planted in Winston\'s consciousness as false hope. The prophetic, mystical register of the sentence (like a Biblical revelation) makes its inversion all the crueller. "We shall meet" implies conspiracy, intimacy, and shared purpose — false solidarity that makes the betrayal total.',
    significance: 'Demonstrates the Party\'s most subtle tactic: they do not suppress hope — they manufacture and exploit it. Winston\'s desire to believe in a resistance is the instrument of his destruction. O\'Brien weaponises Winston\'s idealism against him. The phrase reveals that totalitarianism understands its subjects\' inner lives and uses that understanding against them.',
    context: 'Orwell draws on religious and prophetic language deliberately — "the place where there is no darkness" echoes the Book of Revelation\'s descriptions of heaven. The Party appropriates the spiritual language of liberation to serve totalitarian ends, a technique Orwell observed in both Stalinist and Nazi propaganda, which frequently borrowed religious iconography to manufacture transcendent loyalty.',
    essay: 'O\'Brien\'s prophetic promise "we shall meet in the place where there is no darkness" operates through devastating dramatic irony, luring Winston — and the reader — with the possibility of refuge, only to reveal that the "place of no darkness" is the perpetually lit torture chambers of the Ministry of Love, demonstrating how the Party weaponises hope itself as an instrument of control.'
  },
  {
    n: 6,
    quote: 'Doublethink means the power of holding two contradictory beliefs in one\'s mind simultaneously, and accepting both of them.',
    speaker: 'From "The Theory and Practice of Oligarchical Collectivism," the book attributed to Goldstein, which Winston reads in Part 2. It functions as the regime\'s own theoretical self-analysis.',
    meaning: 'The ability to genuinely believe two opposite things at once — without discomfort — is the psychological mechanism that makes the Party\'s entire system possible.',
    techniques: 'Definition as literary technique — the clinical, encyclopaedic register gives it the air of technical authority, making a monstrous idea sound merely administrative. "Accepting both of them" is the crucial element — not just holding contradictions but genuinely believing both, without distress. Orwell coined "doublethink" as a neologism, demonstrating that language itself is being weaponised.',
    significance: 'Doublethink is the deepest level of mental colonisation — beyond censorship, beyond propaganda, into the complete reconstruction of how the mind processes reality. It explains how citizens can simultaneously know and not know, believe and disbelieve. Without it, citizens could not accept that the Ministry of Peace wages war, or that the Ministry of Love tortures.',
    context: 'Orwell observed doublethink in Soviet communism (Trotsky was simultaneously the revolution\'s hero and its greatest traitor, depending on the year) and in the intellectuals who convinced themselves that Stalinist atrocities were a necessary price for progress. He also saw it in wartime propaganda, which sold mass death as liberation. The concept anticipates cognitive dissonance theory in psychology.',
    essay: 'Orwell\'s definition of "doublethink" — the capacity to hold two contradictory beliefs simultaneously and accept both as true — represents the most sophisticated instrument in the Party\'s arsenal, demonstrating that the ultimate ambition of totalitarianism is not to suppress thought but to colonise it entirely, making logical contradiction feel natural and even comforting.'
  },
  {
    n: 7,
    quote: 'Who controls the past controls the future: who controls the present controls the past.',
    speaker: 'A Party slogan that Winston also reflects upon; it also functions as a theoretical statement of the regime\'s strategy. It directly explains Winston\'s own work at the Ministry of Truth.',
    meaning: 'Whoever rewrites history can shape what people believe about the future; and whoever holds present power controls which version of history is accepted. The logic is circular and inescapable.',
    techniques: 'Chiasmus — the two clauses mirror and reverse each other, creating a circular trap from which there is no logical exit. The colon connects two equally disturbing propositions as equal truths. Repetition of "controls" and "past/future/present" makes the slogan feel totalising — every dimension of time is implicated. The aphoristic structure makes it sound like a profound truth rather than an admission of manipulation.',
    significance: 'Explains Winston\'s job: by constantly rewriting records to match the Party\'s current position, the Party creates a version of the past that makes the present seem inevitable. If there are no records of life before Big Brother, then Big Brother always existed. The quote reveals that the Party\'s most powerful tool is not violence but historiography — the control of collective memory.',
    context: 'Directly reflects Stalinist practice: Trotsky was airbrushed from Soviet photographs after falling out of favour; encyclopaedias were issued with replacement pages to insert over entries about former heroes who had become enemies of the state. Orwell had personally seen newspapers fabricate coverage of the Spanish Civil War and was haunted by what it meant that truth could be retroactively erased.',
    essay: 'The Party slogan "who controls the past controls the future: who controls the present controls the past" reveals through chiasmus that Orwell\'s dystopia is fundamentally a war on memory — by rewriting history, the regime makes the present seem inevitable and the future closed, rendering resistance literally unthinkable because the past from which change might be imagined no longer exists.'
  },
  {
    n: 8,
    quote: 'Until they become conscious they will never rebel, and until after they have rebelled they cannot become conscious.',
    speaker: 'Winston, writing in his diary about the proles in Part 1. This is his recognition of — and despair at — the paradox of collective resistance.',
    meaning: 'The proles can only rebel if they develop political consciousness; but they can only develop consciousness through the experience of rebellion. The two conditions are mutually dependent in a closed loop with no entry point.',
    techniques: 'Paradox — the sentence is structurally a closed loop. Repetition of "conscious" and "rebel" reinforces the circularity. The "until...until" conditional structure creates an impossible equation; both conditions must precede the other. The despairing, philosophical tone implies Winston has recognised a problem he cannot solve. The sentence\'s length and structure mirrors the entrapment it describes.',
    significance: 'Orwell\'s most pessimistic observation about collective resistance. The proles — 85% of Oceania — theoretically have the numbers to overthrow the Party, but lack the political awareness to do so. And the systems that might raise their awareness (education, free press, assembly) are all controlled by the Party. The paradox represents the ultimate form of totalitarian closure: resistance is structurally impossible before it begins.',
    context: 'Reflects Orwell\'s complex and often frustrated relationship with the working class as a political force. He was a socialist who believed in workers\' power but observed that propaganda and material deprivation prevent the working class from recognising its own interests — a concept derived from Marx\'s theory of "false consciousness." His earlier work (The Road to Wigan Pier) explored this tension directly.',
    essay: 'Winston\'s circular observation that the proles cannot rebel without consciousness and cannot gain consciousness without rebelling exposes the most insidious dimension of Orwell\'s totalitarianism: the system has not merely suppressed resistance but engineered a structural paradox in which collective political awakening becomes logically impossible before it can begin.'
  },
  {
    n: 9,
    quote: 'He loved Big Brother.',
    speaker: 'The narrator; the novel\'s final four words. They describe Winston Smith\'s psychological state after his time in Room 101 and the Ministry of Love.',
    meaning: 'Winston Smith, who has spent the entire novel struggling to preserve his identity, rational thought, and love for Julia, has been so completely broken and remade that he now genuinely loves the system that destroyed him.',
    techniques: 'Devastating brevity — four words after hundreds of pages of Winston\'s complex inner life. The third person ("He") creates distance, as if Winston the subject no longer exists as a continuous self. Past tense "loved" is ambiguous — it may suggest finality, or that this is now his permanent state. Capitalisation of "Big Brother" performs the submission the text describes. The absence of exclamation, joy, or self-awareness is the most chilling element.',
    significance: 'The most devastating conclusion Orwell could have written. Winston has not been killed, exiled, or merely silenced — he has been entirely remade. Room 101 does not destroy Winston; it replaces him with a being who is genuinely content in submission. This denies the reader any consolation: there is no martyrdom, no resistance preserved in secret. The self can be utterly and permanently annihilated.',
    context: 'Based partly on what Orwell understood about Soviet show trials, where former Bolshevik heroes confessed enthusiastically to crimes they hadn\'t committed and seemed genuinely converted rather than merely coerced. Orwell was disturbed by the completeness of these transformations. The ending insists that this is possible and that it is the most terrifying political fact of the twentieth century.',
    essay: 'The novel\'s final four words — "He loved Big Brother" — deliver Orwell\'s most devastating conclusion: that the Party\'s ultimate achievement is not the suppression of resistance but its complete inversion, transforming Winston\'s hard-won intellectual defiance into genuine, uncoerced devotion, and extinguishing the possibility of martyrdom or preserved selfhood entirely.'
  },
  {
    n: 10,
    quote: 'Thoughtcrime does not entail death: thoughtcrime IS death.',
    speaker: 'Winston\'s narrated thoughts, Part 1. He is reflecting on the implications of the regime\'s criminalisation of dissenting thought.',
    meaning: 'Thinking forbidden thoughts is not merely punishable by death — it is itself a kind of death, because it destroys your existence as a coherent member of this society and signals the collapse of the self.',
    techniques: 'Antithesis followed by correction — the structure sets up a logical expectation ("does not entail") and violently subverts it ("IS death"). Capitalisation of "IS" performs the certainty and absolutism of the statement. The colon creates a pause that makes the second clause feel like a revelation. Paradox — a thought cannot literally be death; but within this world\'s logic, it is. The compressed syntax enacts the sudden, total nature of the threat.',
    significance: 'Reveals the extent to which the Party has collapsed the distinction between thought and action. In most legal and ethical systems, intent must become action to be punishable. The Party punishes the thought itself — abolishing the boundary between the internal world (which should be inviolable) and the external world (which the state controls). There is no sanctuary left.',
    context: 'Reflects Orwell\'s reading of Soviet purges, where people were arrested for what they might think or who they associated with, not for acts committed. It also anticipates the McCarthyite atmosphere building in the United States — where political beliefs, not actions, were becoming grounds for persecution, career destruction, and social death.',
    essay: 'Winston\'s reflection that "thoughtcrime IS death" — punctuated by deliberate capitalisation — crystallises Orwell\'s vision of a regime that has abolished the last private refuge of selfhood; by criminalising thought itself, the Party makes the interior life as subject to surveillance and punishment as any public act, leaving the individual with no inviolable domain.'
  },
  {
    n: 11,
    quote: 'The Party told you to reject the evidence of your eyes and ears. It was their final, most essential command.',
    speaker: 'Winston\'s narrated thoughts, Part 2. He is articulating what he understands to be the deepest level of the Party\'s demand on its citizens.',
    meaning: 'The ultimate and most important demand the Party makes is not political compliance but sensory denial — citizens must disbelieve what they directly perceive if it contradicts the Party\'s official version of reality.',
    techniques: 'The word "final" implies a hierarchy of commands, with this at the apex — this is the Party\'s core project, not a side effect. "Most essential" reinforces the primacy of this demand above all others. The declarative sentence structure gives it the quality of testimony — as if Winston is reporting evidence. The contrast between "eyes and ears" (direct physical experience) and the Party\'s authority dramatises the violence of the demand.',
    significance: 'Explains why no act of private resistance is ultimately safe: the Party requires not just behavioural compliance but perceptual. Winston cannot even trust his own senses as a basis for dissent. This is the logical culmination of doublethink: the self-evidence of direct experience is insufficient against the Party\'s construction of reality.',
    context: 'Directly echoes totalitarian propaganda theory: Nazi Propaganda Minister Joseph Goebbels understood that if the same lie was repeated often enough, it would be believed over lived experience. Orwell saw this in action in Spain when his own first-hand accounts of events were dismissed in favour of politically convenient fabrications printed in newspapers he had previously trusted.',
    essay: 'Orwell\'s observation that the Party\'s "final, most essential command" was to reject the evidence of one\'s own eyes and ears reveals totalitarianism\'s ultimate ambition: not merely to control behaviour or speech, but to colonise perception itself, making the citizen\'s own sensory experience an unreliable witness against state authority.'
  },
  {
    n: 12,
    quote: 'Reality exists in the human mind, and nowhere else.',
    speaker: 'O\'Brien, during Winston\'s interrogation in Part 3. This is O\'Brien\'s statement of the Party\'s metaphysical position — the philosophical foundation of its total power.',
    meaning: 'O\'Brien argues that there is no objective reality independent of minds — and since the Party controls minds (through doublethink, torture, and psychological remoulding), the Party controls reality itself.',
    techniques: 'Philosophical register — O\'Brien speaks the language of epistemology, framing political domination as a metaphysical position rather than a political one. "And nowhere else" is absolute and admits no exception or loophole. The calm, declarative tone makes the horror of the claim more chilling — this is not said with anger or triumph but with the certainty of established fact. The short sentence structure mirrors the finality of the argument.',
    significance: 'This is the novel\'s philosophical climax. O\'Brien is not making a political argument but an ontological one: if reality is a mental construct, and the Party controls the mental landscape, then the Party does not merely rule the world — it IS the world. This closes off any appeal to objective truth as a basis for resistance. You cannot appeal to reality against the Party because the Party defines reality.',
    context: 'Orwell was deeply troubled by the influence of idealist philosophy on political thought. He saw in Stalinist ideology a similar move: facts were whatever the Party declared them to be, and this was not merely hypocrisy but a genuine epistemological position. O\'Brien\'s statement is Orwell\'s way of forcing readers to see where the logic of "truth is constructed" leads when placed in the hands of authoritarian power.',
    essay: 'O\'Brien\'s chilling declaration that "reality exists in the human mind, and nowhere else" elevates the Party\'s dominance from the political to the metaphysical, arguing that by controlling thought, the Party controls the fabric of existence itself — a position that closes off any appeal to objective truth as a basis for resistance and makes Winston\'s defeat philosophically, not merely physically, inevitable.'
  },
  {
    n: 13,
    quote: 'Nothing was your own except the few cubic centimetres inside your skull.',
    speaker: 'Winston\'s narrated thoughts, Part 1. He is cataloguing what the Party has taken and what — barely — remains.',
    meaning: 'Under the Party\'s rule, the only territory a person genuinely retains is the physical space of their brain — their innermost private thought. Everything else — property, relationships, history, language — has been confiscated.',
    techniques: 'Precision of "cubic centimetres" makes the metaphor physical and biological — the self is measured in tissue. The absolute "nothing" opens the sentence with totalising deprivation. The adjective "few" emphasises how minuscule this last reserve of selfhood actually is. The possessive "your" is deeply ironic given the context — this may not be truly yours either. The sentence\'s quiet tone makes it more disturbing than rhetoric would.',
    significance: 'Defines what is at stake in the novel: the sanctity of private thought as the absolute minimum condition for personhood. It makes Room 101\'s invasion of even this last space — through torture and psychological manipulation — the novel\'s most fundamental act of terror. Orwell is establishing a philosophical threshold: if this is taken, the person is gone.',
    context: 'Written at a time when both Nazi and Soviet regimes had demonstrated through show trials and enhanced interrogation that even the most private beliefs could be extracted and publicly renounced. Orwell\'s point is that if a state can penetrate the inner life through surveillance, torture, and psychological manipulation, then personhood itself — not just freedom — is abolished.',
    essay: 'Winston\'s observation that "nothing was your own except the few cubic centimetres inside your skull" defines the last irreducible boundary of selfhood, making the Party\'s eventual invasion of even this space the novel\'s most fundamental act of terror — the abolition not merely of freedom, but of the very precondition of being a person.'
  },
  {
    n: 14,
    quote: 'Power is not a means; it is an end.',
    speaker: 'O\'Brien, Part 3, during Winston\'s interrogation. This is O\'Brien\'s explicit distinction between the Party and all previous historical regimes.',
    meaning: 'Every previous regime sought power in order to achieve some other goal — prosperity, racial purity, spreading ideology. The Party seeks power purely for its own sake, and this is not a corruption of its purpose but its entire purpose.',
    techniques: 'Antithesis — "means" versus "end" is a classic philosophical distinction, and O\'Brien\'s inversion is logically shocking. Brevity: a single short sentence after which there is nothing to add. The semicolon gives both clauses equal weight, making the contrast emphatic and deliberate. The absolute absence of justification, qualification, or hedging mirrors the absolutism of the claim. The calm, tutorial register (O\'Brien speaking to a student) makes the statement more disturbing.',
    significance: 'The moment Orwell makes explicit what distinguishes his imagined totalitarianism from all previous political systems. The Party has no ideology it is protecting — it has abandoned even the pretence of serving a higher cause. This is Orwell\'s darkest insight: not that tyrants are misguided idealists, but that they may simply enjoy cruelty and domination for their own sake.',
    context: 'A response to the liberal argument that fascism and Stalinism were failed attempts at legitimate political goals. Orwell rejected this: both systems pursued power not as a means but as the point itself. He was influenced by James Burnham\'s theory of the managerial class — a new elite that sought power across both capitalist and communist systems, regardless of ideology.',
    essay: 'O\'Brien\'s declaration that "power is not a means; it is an end" represents the novel\'s most unsettling political thesis — that the Party is not a misguided ideology pursuing a mistaken goal but a pure expression of domination for its own sake, stripping away any moral framework that might allow resistance to be framed as the defence of a higher value.'
  },
  {
    n: 15,
    quote: 'We are the dead.',
    speaker: 'Winston to Julia, in the room above Charrington\'s junk shop in Part 2. Then the hidden telescreen responds: "You are the dead." The moment performs the violation it describes.',
    meaning: 'Winston acknowledges that he and Julia are already effectively dead — their resistance has no future. Then the telescreen\'s repetition reveals that even this private acknowledgement of doom has been heard and will be used against them.',
    techniques: 'The shift from "We are" (Winston\'s intimate, private acknowledgement) to "You are" (the telescreen\'s cold, external correction) performs the collapse of the private/public boundary in real time. Paradox — they are alive but speak of themselves as already dead. The repetition by the telescreen transforms Winston\'s personal honesty into evidence of his guilt. The brevity of both phrases (three and three words) makes the exchange feel like a verdict being delivered and confirmed.',
    significance: 'The most dramatically compressed moment in the novel. Winston has internalised his defeat so completely that he announces it himself — and in doing so, gives the Party the opportunity to confirm it. The transition from private voice to telescreen enacts the collapse of the private sphere that the entire novel has been building toward. The room they believed was an escape was always a trap.',
    context: 'Reflects the practice in Soviet "self-criticism" sessions, where citizens were encouraged to publicly denounce their own thoughts and beliefs. By having Winston speak his own doom, Orwell suggests that totalitarianism ultimately gets its victims to perform their own destruction. The moment also enacts a common pattern of surveillance — the watched subject internalising the watcher\'s perspective until they narrate themselves from the outside.',
    essay: 'When Winston declares "we are the dead" and the hidden telescreen confirms "you are the dead," Orwell dramatises in a single exchange the complete collapse of the private sphere — revealing that the room Winston believed was a sanctuary was always a surveillance apparatus, and that even the most intimate acknowledgement of resistance is immediately converted into evidence of guilt.'
  }
];

// ─── Content: Film Techniques ───────────────────────────────────────────────

const filmTechniques = [
  {
    n: 1,
    name: 'Chiaroscuro / Low-Key Lighting',
    description: 'A cinematographic technique using extreme contrasts between light and shadow, with large portions of the frame left in darkness. Director of Photography Ben Davis deploys this throughout.',
    example: 'V is almost always filmed in deep shadow, with only his mask or the glint of his costume visible. The Norsefire regime is shot in harsh, flat institutional light — interrogation rooms, BTN studios, and the Chancellor\'s chambers are all cold and overexposed.',
    effect: 'The shadow makes V literally ungovernable by the camera\'s gaze — he cannot be fully seen, quantified, or identified, mirroring his resistance to the state\'s surveillance apparatus. The regime\'s harsh lighting feels clinical and oppressive, the visual equivalent of the telescreen in 1984.',
    significance: 'Light = exposure and state control; shadow = freedom and resistance. McTeigue makes the visual language of the film perform its political argument without dialogue. Citizens exist fully in the light — visible, catalogued, known. V operates in the dark — illegible, uncontrollable.',
    essay: 'McTeigue\'s use of chiaroscuro lighting — placing V in constant shadow while the Norsefire regime is bathed in harsh institutional light — visually encodes the film\'s central argument that freedom requires the ability to escape the state\'s gaze, and that the politics of visibility is the politics of power.'
  },
  {
    n: 2,
    name: 'Symbolism — The Guy Fawkes Mask',
    description: 'A visual symbol is an object or image whose meaning extends beyond its literal appearance to carry thematic or ideological weight. The Guy Fawkes mask — white, fixed-smiling, based on the historical figure of the 1605 Gunpowder Plot — is the film\'s central symbol.',
    example: 'The mask is worn by V in every public appearance. In the film\'s climax, hundreds of ordinary citizens don identical masks and march on Parliament. The mask becomes more prominent as the regime weakens.',
    effect: 'The mask does several things simultaneously: it renders V anonymous and therefore unkillable (you cannot kill an idea); its fixed smile creates an uncanny, theatrical effect — V is always performing; when mass-produced and worn by hundreds, it demonstrates that V is not an individual but a symbol anyone can embody. The identical masks make the crowd illegible to the state\'s identifying systems.',
    significance: 'Directly contrasts with 1984\'s surveillance state, where every face must always be exposed and legible to the telescreen. The mask is the visual answer to Big Brother\'s all-seeing eye: where the state demands to see every individual face, the revolution makes all faces identical and therefore unreadable.',
    essay: 'The Guy Fawkes mask functions as the film\'s central visual symbol of collective identity — its anonymity rendering V immune to the state\'s identifying gaze, and its proliferation among ordinary citizens in the climax enacting McTeigue\'s thesis that resistance becomes unstoppable when it transforms from individual rebellion into a shared idea that anyone can carry.'
  },
  {
    n: 3,
    name: 'Musical Motif / Leitmotif — Beethoven\'s Fifth Symphony',
    description: 'A leitmotif is a recurring musical theme associated with a character, idea, or theme. McTeigue uses Beethoven\'s Fifth Symphony ("da-da-da-DUM") as V\'s defining sonic signature.',
    example: 'V plays Beethoven\'s Fifth over the Proms concert speaker system before destroying the Old Bailey. The four-note motif recurs throughout the film, and V uses it as a calling card in his pirate broadcast.',
    effect: 'The Beethoven creates layers of meaning: culturally, the Fifth is associated with triumph over adversity (it was used as a V-for-Victory signal on BBC radio during WWII); mathematically, the four opening notes are Morse code for the letter V; musically, the insistent, unstoppable rhythm mirrors the revolution itself. The choice of classical music — banned by the regime — also makes V\'s act of playing it culturally as well as physically subversive.',
    significance: 'McTeigue uses music as an act of cultural resistance: V restores forbidden art at the same moment as he destroys a symbol of authority. The film argues that the reclamation of suppressed culture is as much a revolutionary act as the destruction of authoritarian architecture.',
    essay: 'McTeigue\'s deployment of Beethoven\'s Fifth Symphony as a recurring leitmotif — whose opening notes encode "V" in wartime Morse code and carry connotations of resistance from WWII broadcasts — positions V\'s revolution as simultaneously aesthetic and political, insisting that the reclamation of suppressed culture is as much an act of defiance as the demolition of authoritarian symbols.'
  },
  {
    n: 4,
    name: 'High Angle / Aerial Shot',
    description: 'A camera position located above the subjects, looking down, which creates effects of scale, smallness, or overwhelming number. Used most significantly in the film\'s climax.',
    example: 'The final sequence, in which the crowd floods the streets toward Parliament, is captured from aerial perspectives, revealing the mass movement from above. Hundreds of Guy Fawkes masks become a visual pattern when seen from height.',
    effect: 'The high angle diminishes individual figures and emphasises the collective mass — the crowd is viewed not as a group of people but as a force of nature. The anonymity of the aerial view means no individual can be identified; from above, the movement is undifferentiated and, crucially, uncontainable.',
    significance: 'This is the visual counterpart to Orwell\'s despairing observation that the proles would never rebel. McTeigue shows what collective awakening looks like from above: not heroic individual figures but an anonymous tide. The same aerial angle that a surveillance state might use to identify and monitor citizens is here repurposed to celebrate their solidarity.',
    essay: 'McTeigue\'s use of high-angle aerial cinematography in the film\'s climax transforms surveillance\'s own visual language into an image of collective liberation — the sea of identical masks viewed from above demonstrating the unstoppable anonymity of a movement too numerous and too undifferentiated for any state apparatus to contain or identify.'
  },
  {
    n: 5,
    name: 'Cross-Cutting / Parallel Editing',
    description: 'An editing technique that alternates between two simultaneous but separate scenes, implying a relationship — contrast, irony, convergence — between them.',
    example: 'McTeigue repeatedly cuts between V\'s methodical preparations (setting the dominoes, arranging the train, broadcasting) and the Norsefire regime\'s increasingly frantic responses — Creedy barking orders, Sutler raging at screens, Finch piecing together evidence.',
    effect: 'The cross-cutting creates sustained tension while simultaneously making the regime look reactive and desperate. The Party is always one step behind: V controls timing and initiative; the regime only responds. This editing structure enacts the power dynamic the film is arguing for, showing who is really in control.',
    significance: 'In stark contrast to 1984, where O\'Brien is always several moves ahead of Winston and the Party\'s omniscience is never genuinely threatened, McTeigue\'s editing places the revolutionary in the position of initiative and mastery. The regime\'s vast infrastructure cannot anticipate or contain a single masked figure with a plan.',
    essay: 'McTeigue\'s cross-cutting between V\'s methodical preparations and the Norsefire regime\'s increasingly frantic responses uses parallel editing to dramatise a complete reversal of power — positioning the lone revolutionary as the controlling, initiating force and the totalitarian government as perpetually reactive, fundamentally undermining the regime\'s projected image of omnipotent control.'
  },
  {
    n: 6,
    name: 'Close-Up — Evey\'s Transformation Scene',
    description: 'An extreme close-up fills the frame with a character\'s face, forcing the viewer into maximum intimacy with their emotional state. McTeigue uses this technique at the film\'s emotional turning point: Evey\'s release from V\'s "prison" and her experience of rain.',
    example: 'When Evey steps outside into the rain after her ordeal, McTeigue uses extreme close-ups of her face — tears mixing with rain — then pulls to a wide shot of her standing in the open. The camera oscillates between intimate and expansive.',
    effect: 'The close-up makes Evey\'s emotional state inescapable — the viewer cannot look away from her face. The combination with the rain (a conventional symbol of baptism, renewal, and emotional release) reinforces that this is a moment of transformation. The transition from close-up to wide shot performs the shift from confinement to freedom — the camera literally expands as she does.',
    significance: 'This is the film\'s emotional and thematic centre — the moment Evey achieves interior freedom. McTeigue\'s camera enacts the transformation rather than merely depicting it, making the viewer feel the shift in scale from confinement to openness. Her journey completes what V began: the creation of a self that the regime cannot reclaim.',
    essay: 'McTeigue\'s sequence of extreme close-ups on Evey\'s face, followed by a wide shot of her in the rain, uses the camera\'s own shift in scale to mirror her psychological transformation — contracting to register her interiority, then expanding to mark her release into a freedom felt in the body before it can be articulated in language.'
  },
  {
    n: 7,
    name: 'Colour Symbolism — Red/Black/Gold vs Grey/Blue',
    description: 'The deliberate selection of a colour palette to associate characters, institutions, and values, allowing ideology to operate through visual impression rather than dialogue.',
    example: 'V\'s costume is black; his world features warm reds, roses, firelight, and gold (the Old Bailey and Parliament explosions). The Norsefire regime is rendered in cold grey, institutional green, and blue — the BTN broadcast sets, interrogation rooms, and Creedy\'s uniforms are all desaturated and cold.',
    effect: 'Red and black are the traditional colours of anarchism, historically representing blood and resistance. Fire imagery (gold/orange) suggests both destruction and warmth. The regime\'s cold palette communicates sterility, bureaucratic control, and emotional hollowness. Audiences absorb this subconsciously: V\'s world feels alive; the regime\'s feels dead.',
    significance: 'Colour does ideological work without a word of dialogue — audiences are primed to experience V\'s revolution as warm and vital and the regime\'s order as cold and lifeless. McTeigue uses the cinematographic palette to embody the political argument, making the film\'s case through sensation as well as narrative.',
    essay: 'McTeigue\'s deliberate colour palette — associating V\'s revolution with warm reds, blacks, and the gold of firelight while rendering Norsefire\'s world in cold institutional grey and blue — allows the film\'s political argument to operate through sensory impression, aligning warmth and vitality with resistance and coldness with the deadening effects of authoritarian control.'
  },
  {
    n: 8,
    name: 'The Domino Sequence — Visual Metaphor and Slow Motion',
    description: 'A visual metaphor uses an image to represent an idea; slow motion photography (undercranking or high-speed capture) stretches time to heighten emotional and aesthetic impact. McTeigue combines both in the film\'s most elaborate single sequence.',
    example: 'V arranges thousands of black and red dominoes on the floor of his gallery in the precise shape of the letter V, then topples the first. The falling is filmed in slow motion. The sequence takes several minutes of screen time.',
    effect: 'Dominoes are a visual metaphor for cascading, irreversible consequence — one action triggers all others in a chain that cannot be reversed once started. The letter V formed by the dominoes makes V\'s own identity the shape of the revolution. Slow motion during the toppling heightens visual beauty and gives the sequence a ceremonial, almost religious quality. The patience of the arrangement contrasts with the chaos of its fall.',
    significance: 'The sequence argues visually that V\'s plan is like the first domino — a revolution set in motion that cannot now be stopped, because each action triggers the next. It also demonstrates that change requires both patient preparation (the careful arrangement) and the willingness to let go of control (the toppling).',
    essay: 'The domino sequence — in which V painstakingly arranges thousands of tiles in the shape of a V, then films their fall in slow motion — functions as the film\'s defining visual metaphor: the revolutionary act requires precise preparation and the understanding that once begun, its consequences cascade beyond any individual\'s control, becoming larger than the person who initiated them.'
  },
  {
    n: 9,
    name: 'Fire and Pyrotechnic Imagery',
    description: 'The repeated, deliberate use of fire, explosions, and fireworks as visual and thematic motifs across multiple key sequences in the film.',
    example: 'The destruction of the Old Bailey (opening act), V\'s burning of Larkhill documents, the Parliament explosion set to the 1812 Overture, and the final fireworks display over London\'s skyline. Each is filmed as spectacle.',
    effect: 'Fire is simultaneously destructive and illuminating — it destroys old structures while making them spectacularly, briefly visible. Explosions staged to the 1812 Overture make destruction operatic and, deliberately, celebratory. Fire carries mythological connotations: Prometheus stole fire from the gods for humanity; fire has long been associated with purification, rebirth, and radical transformation.',
    significance: 'McTeigue makes destruction beautiful, which is deliberately provocative. The explosions are presented not as violence but as art — V is explicitly a performer, and his attacks are theatrical events. This challenges audiences to consider the relationship between spectacle, destruction, and political change: can the obliteration of authoritarian architecture be experienced as liberation?',
    essay: 'McTeigue\'s recurring fire and pyrotechnic imagery — from the Old Bailey\'s destruction to the final explosion of Parliament, each staged as theatrical spectacle set to music — positions V\'s acts of destruction as simultaneously violent and aesthetic, challenging audiences to confront whether the demolition of authoritarian symbols can constitute a form of liberation rather than merely terror.'
  },
  {
    n: 10,
    name: 'Final Tracking Shot — Unmasking of the Crowd',
    description: 'A tracking shot is a camera movement that follows horizontally across a scene. McTeigue uses one in the film\'s final moments to reveal the faces beneath the Guy Fawkes masks as the revolution completes.',
    example: 'After the Parliament explosion, the camera tracks along the unmasked crowd, revealing individual faces — including Evey, and, in the faces of strangers, echoes of Gordon, Delia, Valerie, and others who died. The dead appear to live on in those who carry their memory.',
    effect: 'The move from masked anonymity to unmasked individual faces performs the film\'s resolution — the revolution has succeeded, and it is now safe to be seen again. The resurrection motif (dead characters reflected in the faces of survivors) suggests that ideas and identities survive through those who carry them forward. The tracking movement gives the sequence a final, unhurried sense of witness.',
    significance: 'Answers the question both texts raise about whether the self can survive totalitarianism. McTeigue suggests it can — not because individual selfhood is indestructible, but because ideas and memories are preserved in the people who loved them and choose to carry them forward. This directly opposes Orwell\'s conclusion in Winston\'s annihilation.',
    essay: 'The final tracking shot — revealing individual human faces beneath removed masks, including echoes of those who died carried in the features of the living — enacts McTeigue\'s resolution to the question both texts pose: that identity survives not as individual selfhood but as collective memory, passed forward through those who choose to carry it, making the revolutionary idea truly unkillable.'
  }
];

// ─── Content: V for Vendetta Quotes ─────────────────────────────────────────

const quotesVFV = [
  {
    n: 1,
    quote: 'People should not be afraid of their governments. Governments should be afraid of their people.',
    speaker: 'V, speaking to Evey Hammond early in the film. This is V\'s foundational political statement — an inversion of the existing power relationship between state and citizen.',
    meaning: 'V argues that the natural and legitimate relationship between government and governed is the reverse of what the Norsefire regime has established. Governments derive their authority from the people, and accountability runs downward — from government to citizen, not the other way.',
    techniques: 'Antithesis — the two sentences directly mirror and reverse each other, creating a rhetorical chiasmus. Anaphora — "People should not... Governments should..." reinforces the inversion structurally. The declarative confidence of both sentences gives them the quality of political axioms — not arguments, but self-evident truths. The shift from negative (what should not be) to positive (what should be) moves from critique to aspiration.',
    significance: 'Directly opposes Orwell\'s vision of permanent, total fear. In 1984, O\'Brien expresses no concern whatsoever about popular opinion; citizens are afraid of Big Brother with no expectation of reciprocity. V\'s line declares that the relationship of fear Orwell depicts is not natural or inevitable but imposed — and therefore reversible.',
    context: 'Echoes social contract theory (Locke, Rousseau): governments derive authority from the people and can be legitimately overthrown when they abuse it. McTeigue was filming during the early War on Terror, when critics argued that Western governments were using the politics of fear — of terrorism, of instability — to justify surveillance and curtailed civil liberties. The line is a direct challenge to that climate.',
    essay: 'V\'s declaration that "governments should be afraid of their people" directly inverts the power dynamic Orwell presents in 1984, repositioning fear not as an instrument of authoritarian control but as the rightful relationship of an accountable government to its citizens — and framing the revolution not as destruction but as the restoration of a proper natural order.'
  },
  {
    n: 2,
    quote: 'Beneath this mask there is more than flesh. Beneath this mask there is an idea, Mr. Creedy, and ideas are bulletproof.',
    speaker: 'V, to Peter Creedy, the head of the secret police, during the film\'s climax — as Creedy\'s men shoot V repeatedly.',
    meaning: 'V cannot be fully killed by shooting him because the mask — and the idea it represents — exists independently of his physical body. Killing V the person does not kill V the concept. The idea continues in everyone who wears the mask or believes in what it represents.',
    techniques: 'Repetition of "beneath this mask" creates anaphoric emphasis, each repetition adding a deeper layer of meaning. The shift from "flesh" to "idea" moves from the material to the abstract, from the mortal to the potentially immortal. "Ideas are bulletproof" is a rhetorical compression of the film\'s entire argument into three words. The dramatic context — V speaking while being shot — makes the claim perform itself in real time before the audience\'s eyes.',
    significance: 'This is the film\'s philosophical climax and the most direct contrast with Orwell\'s conclusion. Winston Smith can be destroyed because he is a person; V cannot be fully destroyed because he is also a symbol. The line explains why the mask must pass to Evey — and then to the crowd — the idea continues through whoever wears it.',
    context: 'Reflects the historical reality of political martyrdom: executing a figurehead can strengthen rather than end a movement (as the British discovered with Irish republican leaders, and as the Romans discovered with early Christian martyrs). McTeigue draws on the legacy of Guy Fawkes himself — whose memory has outlasted the state that executed him by four hundred years.',
    essay: 'V\'s climactic assertion that "ideas are bulletproof" — delivered as his body absorbs gunfire — presents the film\'s ultimate counterargument to Orwell\'s vision of individual resistance crushed by state power, insisting that when rebellion transcends the individual and becomes symbolic, it enters a register where physical violence can no longer fully extinguish it.'
  },
  {
    n: 3,
    quote: 'Remember, remember the fifth of November, the Gunpowder Treason and Plot. I know of no reason why the Gunpowder Treason should ever be forgot.',
    speaker: 'V, in the film\'s opening sequence. He speaks the real historical rhyme commemorating Guy Fawkes\'s 1605 plot to blow up the Houses of Parliament — a failed act of resistance that has nonetheless been remembered for four centuries.',
    meaning: 'V invokes a centuries-old tradition of remembering political resistance. The rhyme establishes that the act of remembering — keeping the memory of resistance alive — is itself a political act. The past is not dead; it is a resource.',
    techniques: 'Rhyme scheme gives the lines a ceremonial, folkloric, incantatory quality — the words feel older than any individual speaker. Repetition of "Gunpowder Treason" and "remember/forgot" creates an almost ritualistic effect, as if summoning historical precedent. The archaic register connects V\'s present action to a tradition of resistance predating the current regime. Direct address to the audience positions them as inheritors of this memory.',
    significance: 'Memory is the film\'s primary political weapon — both for the regime (which suppresses it) and for V (who deploys it). By invoking Fawkes, V establishes that resistance has a history, and that history matters. This directly opposes the Party\'s erasure of the past in 1984, where the destruction of memory is the foundation of control.',
    context: 'Guy Fawkes Day (November 5th) is a real British cultural tradition. Moore and McTeigue reframe it from a celebration of the Protestant state\'s survival to a celebration of the idea of principled resistance to unjust authority. The mask and rhyme have since become globally associated with protest movements (the Anonymous movement, Occupy) — demonstrating that the film\'s symbolic vocabulary has become genuinely politically generative.',
    essay: 'V\'s incantation of the Guy Fawkes rhyme — with its folkloric rhyme scheme and archaic register — establishes memory as V for Vendetta\'s central political instrument, positioning the act of historical remembrance as itself a revolutionary practice that resists the totalitarian impulse, so powerfully depicted in 1984, to erase the past and thereby foreclose the future.'
  },
  {
    n: 4,
    quote: 'Artists use lies to tell the truth.',
    speaker: 'V, to Evey, in dialogue about the nature of art, deception, and truth. V has at this point admitted to constructing an elaborate false scenario for Evey\'s psychological transformation.',
    meaning: 'Art — which is not factually literal — can reveal emotional, psychological, or political truths that direct reporting or factual discourse cannot access. The "lie" of fiction or performance can tell deeper truths than the "truth" of documentation.',
    techniques: 'Paradox — "lies" and "truth" are opposites placed in productive, deliberate tension. The brevity of the statement gives it the quality of an aphorism — a compressed wisdom. The context (V confessing to deception while justifying it) makes this a simultaneously self-incriminating and self-defending statement. The word "artists" aligns V with a creative tradition rather than a political or military one.',
    significance: 'This line defends the film\'s own method. V for Vendetta is fiction — a "lie" — but uses that fiction to tell political truths about surveillance, fascism, and resistance that documentary might not be able to access. It also defends all of V\'s cultural acts of resistance (film screenings, music, roses) as serious political interventions, not mere entertainment.',
    context: 'Connects to a long tradition of art as political resistance — from underground protest songs to samizdat literature in the Soviet Union, where banned works were hand-copied and passed in secret. McTeigue, making a film in the aftermath of 9/11 and the Iraq War, was also claiming the right of fiction to critique political reality that mainstream media was being pressured not to address. The film was itself accused of glorifying terrorism — proving the point.',
    essay: 'V\'s paradoxical claim that "artists use lies to tell the truth" positions the creative act as a form of resistance that operates in a register unavailable to propaganda or direct political confrontation, suggesting that fiction — including the film itself — can access and transmit truths about power that the regime\'s control of official information cannot suppress.'
  },
  {
    n: 5,
    quote: 'It is the very last inch of us. But within that inch, we are free.',
    speaker: 'Valerie, a lesbian actress imprisoned and eventually killed by the Norsefire regime; her words are recorded in a letter written on toilet paper and passed secretly between cells, eventually reaching Evey. It is the film\'s most emotionally powerful piece of writing.',
    meaning: 'Even in the most complete totalitarian imprisonment — when every other freedom has been stripped away — the regime cannot take the final, irreducible interior space of the self. And within that last space, however tiny, the individual remains genuinely free.',
    techniques: 'The metaphor of the "inch" makes the last territory of selfhood physically tiny but morally absolute — small in scale, infinite in significance. The contrast between "last inch" (smallness, fragility) and "we are free" (absolute, unqualified) creates a powerful paradox: the smallest space contains the greatest freedom. The first-person voice gives the statement the intimacy and authority of testimony. The handwritten letter\'s physical fragility — written on toilet paper in a prison cell — enacts the vulnerability and determination it describes.',
    significance: 'This is the film\'s most direct response to Orwell\'s conclusion. Winston Smith loses his last inch — the "few cubic centimetres inside your skull" — to O\'Brien\'s techniques in Room 101. Valerie, by contrast, preserves hers even unto death. Her letter argues that if the self chooses not to surrender its integrity, torture and death cannot take the last inch. This is the film\'s counterargument to 1984\'s pessimism: not triumphalist, but insisting on the possibility of a dignity the state cannot confiscate.',
    context: 'Draws on real LGBTQ+ histories of persecution under fascism — gay men were among those imprisoned in Nazi concentration camps, identified by the pink triangle — and under contemporary authoritarian regimes. The letter is the film\'s most direct engagement with the human cost of state persecution, making the Norsefire regime\'s violence specific and historically grounded rather than merely hypothetical.',
    essay: 'Valerie\'s declaration that "within that inch, we are free" offers V for Vendetta\'s most direct and moving response to Orwell\'s vision of total psychological colonisation in 1984 — asserting that the deepest interior space of selfhood cannot be taken by the state if the individual chooses not to surrender it, positioning integrity itself as the last and most durable form of resistance.'
  }
];

// ─── Document Builder ────────────────────────────────────────────────────────

function build1984Section(q, color = '1A5276', qBg = 'D6EAF8') {
  return [
    h2(`Quote ${q.n}: "${q.quote}"`, color),
    quoteBlock(q.quote, qBg, color),
    row('Speaker / Location', q.speaker, color),
    row('Literal Meaning', q.meaning, color),
    row('Literary Techniques', q.techniques, color),
    row('Thematic Significance', q.significance, color),
    row('Historical & Contextual Reference', q.context, color),
    essayBox(q.essay),
    divider()
  ];
}

function buildTechniqueSection(t, color = '922B21', tBg = 'FADBD8') {
  return [
    h2(`Technique ${t.n}: ${t.name}`, color),
    quoteBlock(t.name, tBg, color),
    row('What It Is', t.description, color),
    row('Specific Example from the Film', t.example, color),
    row('Effect Created', t.effect, color),
    row('Thematic Significance', t.significance, color),
    essayBox(t.essay, 'FDEDEC', '922B21'),
    divider('D98880')
  ];
}

function buildVFVQuoteSection(q, color = '6C3483', qBg = 'E8DAEF') {
  return [
    h2(`Quote ${q.n}: "${q.quote.substring(0, 60)}..."`, color),
    quoteBlock(q.quote, qBg, color),
    row('Speaker / When', q.speaker, color),
    row('Literal Meaning', q.meaning, color),
    row('Techniques', q.techniques, color),
    row('Thematic Significance', q.significance, color),
    row('Historical & Contextual Reference', q.context, color),
    essayBox(q.essay, 'F5EEF8', '6C3483'),
    divider('C39BD3')
  ];
}

// ─── Assemble Children ───────────────────────────────────────────────────────

const children = [
  new Paragraph({
    children: [new TextRun({ text: 'Comparative Text Study — Detailed Quote & Technique Reference', bold: true, size: 40, color: '1F4E79', font: 'Arial' })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 120 }
  }),
  new Paragraph({
    children: [new TextRun({ text: 'Nineteen Eighty-Four (Orwell, 1949)  |  V for Vendetta (McTeigue, 2005)', size: 24, color: '555555', font: 'Arial', italics: true })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 360 }
  }),

  // ── Part 1 ──
  h1('PART 1: NINETEEN EIGHTY-FOUR — 15 KEY QUOTES', '1F4E79'),
  new Paragraph({
    children: [new TextRun({ text: 'Each entry includes: speaker/location · literal meaning · literary techniques · thematic significance · historical context · ready-to-use essay sentence.', size: 22, color: '555555', font: 'Arial', italics: true })],
    spacing: { before: 0, after: 240 }
  }),
  ...quotes1984.flatMap(q => build1984Section(q)),

  // ── Part 2 ──
  h1('PART 2: V FOR VENDETTA (FILM) — 10 KEY FILM TECHNIQUES', '7B241C'),
  new Paragraph({
    children: [new TextRun({ text: 'Each entry includes: what it is · specific example from the film · effect created · thematic significance · ready-to-use essay sentence.', size: 22, color: '555555', font: 'Arial', italics: true })],
    spacing: { before: 0, after: 240 }
  }),
  ...filmTechniques.flatMap(t => buildTechniqueSection(t)),

  // ── Part 3 ──
  h1('PART 3: V FOR VENDETTA (FILM) — 5 KEY QUOTES', '4A235A'),
  new Paragraph({
    children: [new TextRun({ text: 'Each entry includes: speaker/when · literal meaning · techniques · thematic significance · historical context · ready-to-use essay sentence.', size: 22, color: '555555', font: 'Arial', italics: true })],
    spacing: { before: 0, after: 240 }
  }),
  ...quotesVFV.flatMap(q => buildVFVQuoteSection(q)),
];

// ─── Create & Save ───────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 22 } }
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 34, bold: true, font: 'Arial' },
        paragraph: { spacing: { before: 400, after: 160 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: 'Arial' },
        paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('1984_VfV_quotes_and_techniques.docx', buf);
  console.log('Done!');
});
